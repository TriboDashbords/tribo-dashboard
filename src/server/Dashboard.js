/**
 * Dashboard.js — дані для сторінки дашборду (агрегація + кеш). Лише ЧИТАЄ Таблицю.
 * Доступ перевіряє Auth.js (вхід за кодом із пошти), тут перевірок немає.
 *
 * Розгортання: Веб-застосунок, "Виконувати як: Користувач, який відкриває застосунок",
 *              "Хто має доступ: Будь-хто з обліковим записом Google".
 */

var DCONFIG = {
  SALES_SHEET: 'sales',
  GROUPS_SHEET: '_groups',
  CONTROL_SHEET: '_control',
  DATASETS_SHEET: '_datasets',
  STALE_HOURS: 2,           // попередження, якщо дані старші
  OTHER_THRESHOLD: 0.05,    // групи з часткою < 5% → "Інше"
  CACHE_SECONDS: 21600,     // 6 год (максимум CacheService); кеш скидається при новому оновленні даних
  TIMEZONE: 'Europe/Kiev'
};

/* ===================== API для сторінки (google.script.run) ===================== */

/** Початкові дані: підсумки року по місяцях + дані поточного місяця. */
function apiInit() {
  var now = new Date();
  var year = parseInt(Utilities.formatDate(now, DCONFIG.TIMEZONE, 'yyyy'), 10);
  var month = parseInt(Utilities.formatDate(now, DCONFIG.TIMEZONE, 'M'), 10);
  var y = getYear_(year);
  return {
    year: year, currentMonth: month,
    months: y.summary.months, ytd: y.summary.total,
    updated: y.summary.updated, staleHours: DCONFIG.STALE_HOURS,
    groupOrder: y.summary.groupOrder, otherThreshold: DCONFIG.OTHER_THRESHOLD,
    period: getPeriodCached_(year, month, y),
    yearPeriod: getPeriodCached_(year, 0, y)
  };
}

/** Дані періоду: period = 1..12 (місяць) або 0 (весь рік). */
function apiPeriod(year, period) {
  return getPeriodCached_(year, period, null);
}

/** Продукти клієнта за період. */
function apiClientItems(year, period, clientKey) {
  var v = dataVersion_();
  var key = 'ci|' + v + '|' + year + '|' + period;
  var all = cacheGetJson_(key);
  if (!all) { buildAndCache_(year, v); all = cacheGetJson_(key) || {}; }
  return all[clientKey] || [];
}

/* ===================== Кеш ===================== */

function getYear_(year) {
  var v = dataVersion_();
  var summary = cacheGetJson_('ys|' + v + '|' + year);
  if (!summary) summary = buildAndCache_(year, v).summary;
  return { summary: summary, version: v };
}

function getPeriodCached_(year, period, y) {
  var v = y ? y.version : dataVersion_();
  var p = cacheGetJson_('p|' + v + '|' + year + '|' + period);
  if (!p) p = buildAndCache_(year, v).periods[periodKey_(year, period)];
  return p || emptyPeriod_();
}

/** Версія даних: час останнього оновлення sales + відбиток аркуша _groups (ручні правки скидають кеш). */
function dataVersion_() {
  var ss = getSpreadsheet_();
  var upd = '';
  var ds = ss.getSheetByName(DCONFIG.DATASETS_SHEET);
  if (ds && ds.getLastRow() > 1) {
    ds.getRange(2, 1, ds.getLastRow() - 1, 5).getDisplayValues().forEach(function (r) { if (r[0] === 'sales') upd = r[4]; });
  }
  var g = ss.getSheetByName(DCONFIG.GROUPS_SHEET), gh = '';
  if (g && g.getLastRow() > 1) {
    var txt = g.getRange(2, 1, g.getLastRow() - 1, 4).getValues().map(function (r) { return r.join('\t'); }).join('\n');
    gh = Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, txt, Utilities.Charset.UTF_8)).slice(0, 10);
  }
  return upd + '#' + gh;
}

/** CacheService: до 100 КБ на значення — великі значення діляться на частини. */
function cachePutJson_(key, obj) {
  var cache = CacheService.getScriptCache();
  var s = JSON.stringify(obj), size = 90000, n = Math.ceil(s.length / size) || 1, items = {};
  for (var i = 0; i < n; i++) items[key + '#' + i] = s.slice(i * size, (i + 1) * size);
  items[key] = String(n);
  cache.putAll(items, DCONFIG.CACHE_SECONDS);
}

function cacheGetJson_(key) {
  var cache = CacheService.getScriptCache();
  var n = cache.get(key);
  if (!n) return null;
  var keys = []; for (var i = 0; i < parseInt(n, 10); i++) keys.push(key + '#' + i);
  var parts = cache.getAll(keys), s = '';
  for (var j = 0; j < keys.length; j++) { if (parts[keys[j]] === undefined) return null; s += parts[keys[j]]; }
  try { return JSON.parse(s); } catch (e) { return null; }
}

/* ===================== Читання Таблиці й агрегація ===================== */

function buildAndCache_(year, version) {
  var ss = getSpreadsheet_();
  var sh = ss.getSheetByName(DCONFIG.SALES_SHEET);
  var rows = [];
  if (sh && sh.getLastRow() > 1) {
    var head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    var vals = sh.getRange(2, 1, sh.getLastRow() - 1, head.length).getValues();
    var idx = {}; head.forEach(function (h, i) { idx[h] = i; });
    var y = String(year);
    vals.forEach(function (r) {
      var m = r[idx.month];
      if (m instanceof Date) m = Utilities.formatDate(m, DCONFIG.TIMEZONE, 'yyyy-MM-dd');
      m = String(m);
      if (m.slice(0, 4) !== y) return;
      rows.push({ month: m, group_path: String(r[idx.group_path]), item_code: String(r[idx.item_code]),
        item: String(r[idx.item]), unit: String(r[idx.unit]), counterparty_code: String(r[idx.counterparty_code]),
        counterparty: String(r[idx.counterparty]), qty: Number(r[idx.qty]) || 0, amount: Number(r[idx.amount]) || 0 });
    });
  }
  var groupsMap = {};
  var g = ss.getSheetByName(DCONFIG.GROUPS_SHEET);
  if (g && g.getLastRow() > 1) g.getRange(2, 1, g.getLastRow() - 1, 4).getValues().forEach(function (r) { groupsMap[String(r[0])] = String(r[3] || r[2] || r[0]); });
  var controls = [];
  var c = ss.getSheetByName(DCONFIG.CONTROL_SHEET);
  if (c && c.getLastRow() > 1) c.getRange(2, 1, c.getLastRow() - 1, 6).getDisplayValues().forEach(function (r) {
    if (r[0] === 'sales') controls.push({ period_from: r[2], period_to: r[3], amount: r[5] === '' ? null : parseFloat(String(r[5]).replace(/[\s ]/g, '').replace(',', '.')) });
  });
  var updated = (version || '').split('#')[0];

  var agg = aggregateYear_(rows, groupsMap, controls, year, updated);
  Object.keys(agg.periods).forEach(function (k) {
    var per = k === String(year) ? 0 : parseInt(k.slice(5, 7), 10);
    cachePutJson_('p|' + version + '|' + year + '|' + per, agg.periods[k]);
    cachePutJson_('ci|' + version + '|' + year + '|' + per, agg.clientItems[k]);
  });
  cachePutJson_('ys|' + version + '|' + year, agg.summary);
  return agg;
}

function periodKey_(year, period) { return period ? year + '-' + ('0' + period).slice(-2) : String(year); }
function emptyPeriod_() { return { total: 0, groupsSum: 0, clientsSum: 0, control: null, rows: 0, groups: [], clients: [] }; }

/**
 * ЧИСТА функція агрегації (без Google API; та сама використовується в попередньому перегляді).
 * rows: [{month:'2026-11-01', group_path, item_code, item, unit, counterparty_code, counterparty, qty, amount}]
 * groupsMap: {group_path: display_name}; controls: [{period_from, period_to, amount}]
 */
function aggregateYear_(rows, groupsMap, controls, year, updated) {
  var periods = {}, clientItems = {}, monthsTot = {};
  function P(k) {
    if (!periods[k]) { periods[k] = { total: 0, rows: 0, g: {}, c: {} }; clientItems[k] = {}; }
    return periods[k];
  }
  var yk = String(year);
  rows.forEach(function (r) {
    var mk = r.month.slice(0, 7);
    var gname = groupsMap[r.group_path] || r.group_path || '(без групи)';
    var ckey = r.counterparty_code || r.counterparty;
    var ikey = (r.item_code || r.item) + '|' + r.unit;
    [mk, yk].forEach(function (k) {
      var p = P(k);
      p.total += r.amount; p.rows++;
      p.g[gname] = (p.g[gname] || 0) + r.amount;
      if (!p.c[ckey]) p.c[ckey] = { key: ckey, name: r.counterparty, amount: 0 };
      p.c[ckey].amount += r.amount;
      var ci = clientItems[k][ckey] || (clientItems[k][ckey] = {});
      if (!ci[ikey]) ci[ikey] = { item: r.item, unit: r.unit, qty: 0, amount: 0 };
      ci[ikey].qty += r.qty; ci[ikey].amount += r.amount;
    });
    monthsTot[mk] = (monthsTot[mk] || 0) + r.amount;
  });

  // контрольні суми 1С за місяць; за рік — сума місячних, якщо є всі
  function controlFor(k) {
    if (k === yk) {
      var s = 0, all = true, any = false;
      Object.keys(monthsTot).forEach(function (mk) { var cm = controlFor(mk); if (cm === null) all = false; else { s += cm; any = true; } });
      return any && all ? r2_(s) : null;
    }
    for (var i = 0; i < controls.length; i++) {
      var c = controls[i];
      if (c.period_from.slice(0, 7) === k && c.period_to.slice(0, 7) === k) return c.amount;
    }
    return null;
  }

  var out = {}, outCI = {};
  Object.keys(periods).forEach(function (k) {
    var p = periods[k];
    var groups = Object.keys(p.g).map(function (n) { return { name: n, amount: r2_(p.g[n]) }; })
      .sort(function (a, b) { return b.amount - a.amount; });
    var clients = Object.keys(p.c).map(function (x) { var c = p.c[x]; return { key: c.key, name: c.name, amount: r2_(c.amount) }; })
      .sort(function (a, b) { return b.amount - a.amount; });
    out[k] = {
      total: r2_(p.total), rows: p.rows,
      groupsSum: r2_(groups.reduce(function (s, g) { return s + g.amount; }, 0)),
      clientsSum: r2_(clients.reduce(function (s, c) { return s + c.amount; }, 0)),
      control: controlFor(k), groups: groups, clients: clients
    };
    outCI[k] = {};
    Object.keys(clientItems[k]).forEach(function (ck) {
      var m = clientItems[k][ck];
      outCI[k][ck] = Object.keys(m).map(function (ik) { var it = m[ik]; return { item: it.item, unit: it.unit, qty: Math.round(it.qty * 1000) / 1000, amount: r2_(it.amount) }; })
        .sort(function (a, b) { return b.amount - a.amount; });
    });
  });

  var yearP = out[yk] || emptyPeriod_();
  var months = [];
  for (var m = 1; m <= 12; m++) {
    var mk = yk + '-' + ('0' + m).slice(-2);
    months.push({ m: m, total: out[mk] ? out[mk].total : 0 });
  }
  return {
    summary: { total: yearP.total, months: months, updated: updated,
      groupOrder: yearP.groups.map(function (g) { return g.name; }) }, // стабільні кольори груп
    periods: out, clientItems: outCI
  };
}

function r2_(x) { return Math.round(x * 100) / 100; }

