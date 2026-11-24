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
      if (/^организация/.test(v) && c[j + 1]) org = String(c[j + 1]).trim();
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

/** Дані вкладки за місяць 'YYYY-MM'. */
function apiPlanFact(month) {
  var v = dataVersion_(), key = 'pf|' + v + '|' + month;
  var c = cacheGetJson_(key); if (c) return c;
  var items = readDataset_('planfact', function (o) { return o.month.slice(0, 7) === month; });
  var clients = readDataset_('planfact_clients', function (o) { return o.month.slice(0, 7) === month; });
  var view = buildPlanFact_(items, clients);
  view.month = month;
  view.control = null;
  var ctl = getSpreadsheet_().getSheetByName(CONFIG.SHEET_CONTROL);
  if (ctl && ctl.getLastRow() > 1) ctl.getRange(2, 1, ctl.getLastRow() - 1, 6).getDisplayValues().forEach(function (r) {
    if (r[0] === 'planfact' && r[2].slice(0, 7) === month) view.control = parseFloat(String(r[5]).replace(/[\s ]/g, '').replace(',', '.'));
  });
  cachePutJson_(key, view);
  return view;
}

/**
 * ЧИСТА функція: рядки товарів і клієнтів → {totals, tree, clients}.
 * Скорочені ключі: n назва, p план, f факт, c собівартість факт, pr прибуток факт, pp прибуток план,
 * qp/qf кількість план/факт, k код, ch діти, it товари.
 */
function buildPlanFact_(items, clientRows) {
  function acc(o, r) { o.p += r.sum_plan; o.f += r.sum_fact; o.c += r.cost_fact; o.pr += r.profit_fact; o.pp += r.profit_plan; o.cp += r.cost_plan; }
  function blank(n) { return { n: n, p: 0, f: 0, c: 0, pr: 0, pp: 0, cp: 0 }; }
  function fin(o) { ['p', 'f', 'c', 'pr', 'pp', 'cp'].forEach(function (k) { o[k] = Math.round(o[k]); }); return o; } // цілі гривні — менший обсяг для телефона
  var byF = function (a, b) { return b.f - a.f || b.p - a.p; };

  var totals = blank('Усього');
  var root = { node: blank(''), ch: {}, it: {} };
  var cl = {};
  items.forEach(function (r) {
    acc(totals, r);
    // дерево груп (без напрямку і клієнта), товари агреговані по всіх клієнтах
    var parts = r.group_path ? r.group_path.split(GROUP_RULE.PATH_SEPARATOR) : ['(без групи)'];
    var cur = root; acc(cur.node, r);
    parts.forEach(function (name) {
      if (!cur.ch[name]) cur.ch[name] = { node: blank(name), ch: {}, it: {} };
      cur = cur.ch[name]; acc(cur.node, r);
    });
    var ik = r.item_code || r.item;
    if (!cur.it[ik]) { cur.it[ik] = blank(r.item); cur.it[ik].k = r.item_code; cur.it[ik].qp = 0; cur.it[ik].qf = 0; }
    acc(cur.it[ik], r); cur.it[ik].qp += r.qty_plan; cur.it[ik].qf += r.qty_fact;
    // клієнти
    var ck = r.counterparty_code || r.counterparty;
    if (!cl[ck]) { cl[ck] = blank(r.counterparty); cl[ck].rg = r.region; cl[ck].it = {}; cl[ck].pay = 0; cl[ck].debt = 0; }
    acc(cl[ck], r);
    var ci = cl[ck].it;
    if (!ci[ik]) { ci[ik] = blank(r.item); ci[ik].k = r.item_code; ci[ik].g = parts[parts.length - 1]; ci[ik].qp = 0; ci[ik].qf = 0; }
    acc(ci[ik], r); ci[ik].qp += r.qty_plan; ci[ik].qf += r.qty_fact;
  });
  (clientRows || []).forEach(function (r) {
    var ck = r.counterparty_code || r.counterparty;
    if (!cl[ck]) { cl[ck] = blank(r.counterparty); cl[ck].rg = r.region; cl[ck].it = {}; cl[ck].pay = 0; cl[ck].debt = 0; }
    cl[ck].pay += r.payments; cl[ck].debt += r.debt;
  });

  function itemsOut(m) {
    return Object.keys(m).map(function (k) { var o = fin(m[k]); o.qp = Math.round(o.qp * 1000) / 1000; o.qf = Math.round(o.qf * 1000) / 1000; return o; }).sort(byF);
  }
  // дерево: ланцюжки з одного підвузла зливаються в одну назву "A › B"
  function out(t) {
    var keys = Object.keys(t.ch), its = Object.keys(t.it);
    while (keys.length === 1 && !its.length) {
      var only = t.ch[keys[0]];
      only.node.n = (t.node.n ? t.node.n + ' › ' : '') + only.node.n;
      t = only; keys = Object.keys(t.ch); its = Object.keys(t.it);
    }
    var o = fin(t.node);
    o.ch = keys.map(function (k) { return out(t.ch[k]); }).sort(byF);
    o.it = itemsOut(t.it);
    return o;
  }
  var tree = out(root);
  var clients = Object.keys(cl).map(function (k) {
    var o = fin(cl[k]); o.key = k; o.pay = Math.round(o.pay * 100) / 100; o.debt = Math.round(o.debt * 100) / 100; o.it = itemsOut(o.it); return o;
  }).sort(byF);
  return { totals: fin(totals), tree: tree, clients: clients };
}
