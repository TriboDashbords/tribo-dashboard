/**
 * PlanFact.js — звіт 1С «План-факт продажів»: розбір Excel і дані для вкладки «План-факт».
 *
 * Структура Excel (перевірено на звіті за 11.2026):
 *   шапка: "Период с: 01.11.2026 по: 30.11.2026", "Организация : …";
 *   рядок заголовків: "Количество (план)", "Количество (факт)", "Сумма (план продажи)", "Сумма (факт продажи)", …;
 *   ієрархія (рівні групування рядків): 0 — напрямок (Украина / Экспорт), 1 — клієнт,
 *   далі групи номенклатури; рядок товару має числовий код у стовпці "Код" і назву в "Номенклатура".
 *   Excel дає максимум 7 рівнів, тому глибокі групи й товари можуть стояти на одному рівні 7 —
 *   товар відрізняємо за кодом, а не за рівнем.
 *   Оплати й борг заповнені лише на рядку клієнта.
 *   Рентабельність = прибуток / собівартість (як у 1С).
 */

var PF_HEADERS = {
  'количество (план)': 'qty_plan', 'количество (факт)': 'qty_fact',
  'сумма (план продажи)': 'sum_plan', 'сумма (факт продажи)': 'sum_fact', 'сумма (факт продаж с ндс)': 'sum_fact_vat',
  'сумма (приход проплат)': 'payments', 'сумма (долг)': 'debt',
  'сумма себестоимости (план)': 'cost_plan', 'сумма себестоимости (факт)': 'cost_fact',
  'прибыль (план)': 'profit_plan', 'прибыль (факт)': 'profit_fact'
};

function pfHeaderIndex_(rows) {
  for (var i = 0; i < rows.length && i < 40; i++) {
    var hits = 0;
    rows[i].cells.forEach(function (v) { if (typeof v === 'string' && PF_HEADERS[v.trim().toLowerCase()]) hits++; });
    if (hits >= 4) return i;
  }
  return -1;
}

function isPlanFactTable_(rows) { return pfHeaderIndex_(rows) >= 0; }

function pfDate_(s) {
  var m = /(\d{2})\.(\d{2})\.(\d{4})/.exec(String(s || ''));
  return m ? m[3] + '-' + m[2] + '-' + m[1] : null;
}

/** Розбір таблиці → {payloads: [planfact, planfact_clients], errors, warnings} */
function parsePlanFactTable_(rows, opts) {
  var errors = [], warnings = [];
  var h = pfHeaderIndex_(rows);
  if (h < 0) return { payloads: [], errors: ['не знайдено заголовків план-факту'], warnings: warnings };

  // шапка: період і організація
  var from = null, to = null, org = opts.organization;
  for (var i = 0; i < h; i++) {
    var c = rows[i].cells;
    for (var j = 0; j < c.length; j++) {
      var v = String(c[j] === undefined ? '' : c[j]).trim().toLowerCase();
      if (/^период с/.test(v)) from = pfDate_(c[j + 1]);
      if (/^по:?$/.test(v)) to = pfDate_(c[j + 1]);
      if (/^организация/.test(v) && c[j + 1] && orgKey_(c[j + 1]) !== orgKey_(org)) org = String(c[j + 1]).trim(); // інша організація
    }
  }
  if (!from || !to) errors.push('не знайдено період ("Период с … по …")');
  if (from && to && from.slice(0, 7) !== to.slice(0, 7)) warnings.push('звіт охоплює кілька місяців (' + from + '…' + to + ') — записано як ' + from.slice(0, 7));

  // стовпці
  var col = {}, hdr = rows[h].cells;
  hdr.forEach(function (v, idx) { if (typeof v === 'string' && PF_HEADERS[v.trim().toLowerCase()]) col[PF_HEADERS[v.trim().toLowerCase()]] = idx; });
  var codeCol = 1, nameCol = 2, labelCol = 1;
  for (var r0 = h + 1; r0 < Math.min(rows.length, h + 4); r0++) {
    rows[r0].cells.forEach(function (v, idx) {
      var t = String(v || '').trim().toLowerCase();
      if (t === 'код') codeCol = idx;
      if (t === 'номенклатура') nameCol = idx;
    });
  }
  if (col.sum_fact === undefined) errors.push('немає стовпця "Сумма (факт продажи)"');
  if (errors.length) return { payloads: [], errors: errors, warnings: warnings };

  var month = from.slice(0, 7) + '-01';
  var num = function (cells, f) { return col[f] === undefined ? 0 : toNumber_(cells[col[f]]); };
  var items = [], clients = [], stack = [], total = null;
  for (var k = h + 1; k < rows.length; k++) {
    var row = rows[k], cells = row.cells;
    var label = cells[labelCol];
    if (label === undefined || label === null || String(label).trim() === '') continue;
    var lab = String(label).trim().replace(/,\s*$/, '');
    if (/^общий итог|^итог/i.test(lab)) { total = cells; continue; }
    if (/^(клонтрагент|контрагент|код|валюта)$/i.test(lab) && row.level === 0 && k <= h + 3) continue; // підзаголовки
    var code = cells[codeCol];
    var isItem = (typeof code === 'number') || /^\d+$/.test(String(code).trim());
    if (!isItem) {
      stack.length = row.level;
      stack[row.level] = lab;
      if (row.level === 1) {
        clients.push({ month: month, organization: org, region: stack[0] || '', counterparty_code: '', counterparty: lab,
          payments: r2pf_(num(cells, 'payments')), debt: r2pf_(num(cells, 'debt')) });
      }
      continue;
    }
    // товар: батьківські групи — усе зі стеку до його рівня (для рівня 7 — включно з групою рівня 7)
    var depth = row.level >= 7 && stack[7] !== undefined ? 8 : row.level;
    var path = stack.slice(0, depth).filter(function (x) { return x !== undefined; });
    if (path.length < 2) { warnings.push('товар без клієнта: ' + cells[nameCol]); continue; }
    items.push({
      month: month, organization: org, region: path[0], counterparty_code: '', counterparty: path[1],
      group_path: path.slice(2).join(GROUP_RULE.PATH_SEPARATOR), group_ids: '',
      item_code: String(code).trim(), item: String(cells[nameCol] || '').trim(),
      qty_plan: r3pf_(num(cells, 'qty_plan')), qty_fact: r3pf_(num(cells, 'qty_fact')),
      sum_plan: r2pf_(num(cells, 'sum_plan')), sum_fact: r2pf_(num(cells, 'sum_fact')), sum_fact_vat: r2pf_(num(cells, 'sum_fact_vat')),
      cost_plan: r2pf_(num(cells, 'cost_plan')), cost_fact: r2pf_(num(cells, 'cost_fact')),
      profit_plan: r2pf_(num(cells, 'profit_plan')), profit_fact: r2pf_(num(cells, 'profit_fact'))
    });
  }
  if (!items.length) errors.push('не знайдено рядків товарів');
  var sumFact = items.reduce(function (s, x) { return s + x.sum_fact; }, 0);
  var control = total ? r2pf_(toNumber_(total[col.sum_fact])) : r2pf_(sumFact);
  if (!total) warnings.push('немає рядка "Общий итог" — контрольна сума не перевіряється');
  if (total && col.sum_plan !== undefined) {
    var sp = items.reduce(function (s, x) { return s + x.sum_plan; }, 0);
    if (Math.abs(sp - toNumber_(total[col.sum_plan])) > 1) warnings.push('план: сума рядків ' + r2pf_(sp) + ' ≠ "Общий итог" ' + toNumber_(total[col.sum_plan]));
  }
  var lastDay = new Date(Date.UTC(+month.slice(0, 4), +month.slice(5, 7), 0)).getUTCDate();
  var base = { organization: org, period_from: month, period_to: month.slice(0, 8) + ('0' + lastDay).slice(-2), mode: 'replace_period' };
  var pf = merge_(base, { dataset: 'planfact', control: { rows: items.length, amount: control }, rows: items });
  var payTotal = r2pf_(clients.reduce(function (s, x) { return s + x.payments; }, 0));
  var pc = merge_(base, { dataset: 'planfact_clients', control: { rows: clients.length, amount: payTotal }, rows: clients });
  return { payloads: [pf, pc], errors: errors, warnings: warnings, period: { from: from, to: to } };
}

function r2pf_(x) { return Math.round((x || 0) * 100) / 100; }
function r3pf_(x) { return Math.round((x || 0) * 1000) / 1000; }

/* ===================== Дані для вкладки «План-факт» ===================== */

function readDataset_(name, filterFn) {
  var ds = DATASETS[name], sh = getSpreadsheet_().getSheetByName(ds.sheet), out = [];
  if (!sh || sh.getLastRow() < 2) return out;
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, ds.columns.length).getValues();
  vals.forEach(function (r) {
    var o = {};
    ds.columns.forEach(function (c, i) {
      var v = r[i];
      if (ds.types[c] === 'number') v = Number(v) || 0;
      else if (v instanceof Date) v = Utilities.formatDate(v, DCONFIG.TIMEZONE, 'yyyy-MM-dd');
      else v = String(v);
      o[c] = v;
    });
    if (!filterFn || filterFn(o)) out.push(o);
  });
  return out;
}

/** Місяці, за які є план-факт (нові першими). */
function apiPlanFactMonths() {
  var v = dataVersion_(), key = 'pfm|' + v;
  var c = cacheGetJson_(key); if (c) return c;
  var sh = getSpreadsheet_().getSheetByName(DATASETS.planfact.sheet), set = {};
  if (sh && sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 1).getDisplayValues().forEach(function (r) { if (r[0]) set[String(r[0]).slice(0, 7)] = 1; });
  var out = Object.keys(set).sort().reverse();
  cachePutJson_(key, out);
  return out;
}

/** Дані вкладки за місяць 'YYYY-MM' (лише ФАКТ). */
function apiPlanFact(month) {
  var v = dataVersion_(), key = 'pf2|' + v + '|' + month;
  var c = cacheGetJson_(key); if (c) return c;
  var ss = getSpreadsheet_();
  var items = readDataset_('planfact', function (o) { return o.month.slice(0, 7) === month; });
  var clients = readDataset_('planfact_clients', function (o) { return o.month.slice(0, 7) === month; });
  // назви продуктових груп — як на вкладці «Продажі» (аркуш _groups + правило 9.3)
  var gmap = {};
  var g = ss.getSheetByName(CONFIG.SHEET_GROUPS);
  if (g && g.getLastRow() > 1) g.getRange(2, 1, g.getLastRow() - 1, 4).getValues().forEach(function (r) { gmap[String(r[0])] = String(r[3] || ''); });
  // одиниці виміру — з набору «Продажі» (у звіті план-факт їх немає)
  var units = {};
  var sh = ss.getSheetByName(DATASETS.sales.sheet);
  if (sh && sh.getLastRow() > 1) {
    var ic = DATASETS.sales.columns.indexOf('item') + 1, uc = DATASETS.sales.columns.indexOf('unit') + 1;
    var iv = sh.getRange(2, ic, sh.getLastRow() - 1, 1).getValues(), uv = sh.getRange(2, uc, sh.getLastRow() - 1, 1).getValues();
    iv.forEach(function (r, i) { if (r[0] && uv[i][0]) units[String(r[0]).trim().toLowerCase()] = String(uv[i][0]); });
  }
  var view = buildPlanFact_(items, clients, gmap, units);
  view.month = month;
  view.control = null;
  var ctl = ss.getSheetByName(CONFIG.SHEET_CONTROL);
  if (ctl && ctl.getLastRow() > 1) ctl.getRange(2, 1, ctl.getLastRow() - 1, 8).getDisplayValues().forEach(function (r) {
    if (r[0] === 'planfact' && r[2].slice(0, 7) === month) { view.control = parseFloat(String(r[5]).replace(/[\s\u00a0]/g, '').replace(',', '.')); view.loaded = r[7]; }
  });
  view.groupOrder = view.groups.map(function (x) { return x.n; });
  cachePutJson_(key, view);
  return view;
}

/** Назва продуктової групи для шляху груп. */
function pfGroupName_(path, gmap) {
  if (gmap[path]) return gmap[path];
  if (!path) return '(без групи)';
  var res = resolveProductGroup_(path);
  return GROUP_RULE.DISPLAY_OVERRIDES[res.name] || res.name;
}

/**
 * ЧИСТА функція: рядки товарів і клієнтів → {totals, groups, clients}. Лише факт.
 * Ключі: n назва, f сума факт (грн без ПДВ), pr маржинальний прибуток, c собівартість, q кількість, u одиниця,
 * it продукти, g групи (у клієнта), debt борг.
 */
function buildPlanFact_(items, clientRows, gmap, units) {
  gmap = gmap || {}; units = units || {};
  function blank(n) { return { n: n, f: 0, pr: 0, c: 0 }; }
  function acc(o, r) { o.f += r.sum_fact; o.pr += r.profit_fact; o.c += r.cost_fact; }
  function fin(o) { o.f = Math.round(o.f); o.pr = Math.round(o.pr); o.c = Math.round(o.c); return o; }
  var byF = function (a, b) { return b.f - a.f || b.q - a.q; };
  function addItem(map, r) {
    var k = r.item_code || r.item;
    if (!map[k]) { map[k] = blank(r.item); map[k].q = 0; map[k].u = units[String(r.item).trim().toLowerCase()] || ''; }
    acc(map[k], r); map[k].q += r.qty_fact;
  }
  function itemsOut(map) {
    return Object.keys(map).map(function (k) {
      var o = fin(map[k]);
      if (/^тис\.?\s*шт/i.test(o.u)) { o.q = o.q * 1000; o.u = 'шт'; } // показуємо в штуках
      o.q = Math.round(o.q * 1000) / 1000; return o;
    })
      .filter(function (o) { return o.f || o.q; }).sort(byF);
  }
  var totals = blank('Усього'), groups = {}, cl = {};
  items.forEach(function (r) {
    if (!r.sum_fact && !r.qty_fact && !r.profit_fact) return; // лише фактичні продажі
    acc(totals, r);
    var gn = pfGroupName_(r.group_path, gmap);
    if (!groups[gn]) { groups[gn] = blank(gn); groups[gn].it = {}; }
    acc(groups[gn], r); addItem(groups[gn].it, r);
    var ck = r.counterparty_code || r.counterparty;
    if (!cl[ck]) { cl[ck] = blank(r.counterparty); cl[ck].rg = r.region; cl[ck].g = {}; cl[ck].debt = 0; }
    acc(cl[ck], r);
    var cg = cl[ck].g;
    if (!cg[gn]) { cg[gn] = blank(gn); cg[gn].it = {}; }
    acc(cg[gn], r); addItem(cg[gn].it, r);
  });
  (clientRows || []).forEach(function (r) {
    var ck = r.counterparty_code || r.counterparty;
    if (cl[ck]) cl[ck].debt += r.debt;
  });
  var gOut = Object.keys(groups).map(function (k) { var o = fin(groups[k]); o.it = itemsOut(o.it); return o; })
    .filter(function (o) { return o.f; }).sort(byF);
  var cOut = Object.keys(cl).map(function (k) {
    var o = fin(cl[k]); o.key = k; o.debt = Math.round(o.debt);
    o.g = Object.keys(o.g).map(function (gk) { var x = fin(o.g[gk]); x.it = itemsOut(x.it); return x; }).sort(byF);
    return o;
  }).filter(function (o) { return o.f; }).sort(byF);
  return { totals: fin(totals), groups: gOut, clients: cOut };
}
