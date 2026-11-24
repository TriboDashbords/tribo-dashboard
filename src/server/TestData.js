/**
 * TestData.js — ВИГАДАНІ тестові дані у форматі 5.1 (суми й клієнти умовні, не з 1С).
 * Використовуються для перевірки коду (tests/run.js) і, за потреби, для завантаження в порожню Таблицю.
 */

var ETALON = {
  year: 2026,
  yearTotal: 104500.0, oct: 12000.0, nov: 2500.0,
  shares: { 'МПС': 45.0, '(2.10) Комплект колодок гальмівних': 20.0 },
  topClient: 'ТОВ "Ключовий клієнт (тест)"', topClientAmount: 25000
};

function generateTestPayloads_() {
  var ORG = 'ТОВ "Білоцерківський завод "ТРІБО""';
  var seed = 20261111;
  function rnd() { // mulberry32 — детермінований генератор
    seed |= 0; seed = seed + 0x6D2B79F5 | 0;
    var t = Math.imul(seed ^ seed >>> 15, 1 | seed);
    t = t + Math.imul(t ^ t >>> 7, 61 | t) ^ t;
    return ((t ^ t >>> 14) >>> 0) / 4294967296;
  }
  function r2(x) { return Math.round(x * 100) / 100; }
  function gid(s) { var h = 0; for (var i = 0; i < s.length; i++) h = (h * 31 + s.charCodeAt(i)) >>> 0; return ('00000000' + h.toString(16)).slice(-8) + '-tst'; }
  var GP = ' > ';

  var CZV = 'ТОВ "Ключовий клієнт (тест)"';
  var C = function (n) { return 'ТОВ "Клієнт-тест ' + (n < 10 ? '0' : '') + n + '"'; };
  var groups = [
    { path: 'Готова продукція > (2) Колодки гальмові > (2.1) Колодки гальмові композиційні > (2.1.1) Колодки гальмові МШС', share: 0.45,
      items: [['2ТР-155', 'тис.шт', 1650000, 'Колодки МШС тип 2ТР'], ['2ТР-180', 'тис.шт', 1720000, 'Колодки МШС тип 2ТР'], ['ТИП-А 10', 'тис.шт', 1580000, '']],
      clients: [CZV, C(29), C(30), C(1), C(2), C(3)] },
    { path: 'Готова продукція > (2) Колодки гальмові > (2.10) Комплект колодок гальмівних', share: 0.20,
      items: [['Комплект КГ-4', 'шт', 5200, ''], ['Комплект КГ-8', 'шт', 9800, '']], clients: [C(4), C(5), C(6), C(7), C(8)] },
    { path: 'Готова продукція > (2) Колодки гальмові > (2.2) Колодки залізничні > (2.1.1) Колодки гальмові залізничні', share: 0.048,
      items: [['КЗ-25', 'тис.шт', 1400000, '']], clients: [C(9), C(10), C(11)] },
    { path: 'Готова продукція > (4) Промислове обладнання > (4.1) Бурове обладнання > (2.1.1) Колодки гальмові для бурових установок', share: 0.042,
      items: [['КБУ-1', 'шт', 3100, ''], ['КБУ-2', 'шт', 4200, '']], clients: [C(12), C(13)] },
    { path: 'Готова продукція > (3) Накладки > (3.1) Накладки фрикційні > (3.1.2) Накладки гальмові > (3.1.2.1) Накладки TEMAC', share: 0.049,
      items: [['Накладка TEMAC 200', 'пог.м', 850, ''], ['Накладка TEMAC 300', 'пог.м', 1100, '']], clients: [C(14), C(15), C(16)] },
    { path: 'Готова продукція > Дафмі > Колодки дискового гальма > Легкові авто', share: 0.046,
      items: [['Дафмі D-101', 'шт', 420, ''], ['Дафмі D-205', 'шт', 510, '']], clients: [C(17), C(18), C(19), C(20)] },
    { path: 'Готова продукція > Дафмі > Колодки барабанного гальма', share: 0.044,
      items: [['Дафмі B-12', 'шт', 380, '']], clients: [C(17), C(21), C(22)] },
    { path: 'Метали', share: 0.049,
      items: [['Круг 56 ст.20', 'кг', 46, ''], ['Лист 4 мм ст.3', 'кг', 39, '']], clients: [C(23), C(24)] },
    { path: 'Готова продукція > (2) Колодки гальмові > (2.3) Вставки гальмівні', share: 0.036,
      items: [['Вставка ВГ-3', 'шт', 260, '']], clients: [C(1), C(25)] },
    { path: 'Готова продукція > Армастар > Диски зчеплення', share: 0.036,
      items: [['Армастар ДЗ-240', 'шт', 1900, '']], clients: [C(26), C(27), C(28)] }
  ];
  // місячні підсумки, грн
  var w = [0.93, 0.97, 1.06, 1.01, 1.04, 0.96, 1.02, 0.98, 1.03], ws = w.reduce(function (a, b) { return a + b; }, 0);
  var monthTot = [], acc = 0;
  for (var i = 0; i < 9; i++) { monthTot[i] = i < 8 ? r2(90000000 * w[i] / ws) : r2(90000000 - acc); acc += monthTot[i]; }
  monthTot.push(12000000, 2500000.55);
  var grand = monthTot.reduce(function (a, b) { return a + b; }, 0);
  var gTot = groups.map(function (g) { return g.share * grand; });

  // матриця місяць × група: шум + IPF під підсумки рядків і стовпців
  var M = monthTot.map(function (mt) { return groups.map(function (g) { return mt * g.share * (0.75 + 0.5 * rnd()); }); });
  for (var it = 0; it < 200; it++) {
    M.forEach(function (row, m) { var s = row.reduce(function (a, b) { return a + b; }, 0); row.forEach(function (v, g) { row[g] = v * monthTot[m] / s; }); });
    groups.forEach(function (g, gi) { var s = 0; M.forEach(function (row) { s += row[gi]; }); M.forEach(function (row) { row[gi] *= gTot[gi] / s; }); });
  }
  M.forEach(function (row, m) {
    var s = 0; row.forEach(function (v, g) { row[g] = r2(v); s += row[g]; });
    row[0] = r2(row[0] + monthTot[m] - s); // точний підсумок місяця
  });

  var czvTarget = 25000000, mpsYear = 0; M.forEach(function (row) { mpsYear += row[0]; });
  var czvFrac = czvTarget / mpsYear, czvAcc = 0;

  var payloads = [];
  for (var m = 0; m < 11; m++) {
    var month = '2026-' + ('0' + (m + 1)).slice(-2) + '-01';
    var rows = [];
    groups.forEach(function (g, gi) {
      var cell = M[m][gi];
      var pairs = [];
      g.clients.forEach(function (cl) { g.items.forEach(function (itm) { if (cl === CZV || rnd() < 0.7) pairs.push([cl, itm]); }); });
      if (!pairs.length) pairs.push([g.clients[0], g.items[0]]);
      var czvPairs = pairs.filter(function (p) { return p[0] === CZV; });
      var others = pairs.filter(function (p) { return p[0] !== CZV; });
      var parts = [];
      var czvAmt = 0;
      if (gi === 0) {
        czvAmt = m === 10 ? r2(czvTarget - czvAcc) : r2(cell * czvFrac);
        czvAcc = r2(czvAcc + czvAmt);
        var cw = czvPairs.map(function () { return 0.5 + rnd(); }), cws = cw.reduce(function (a, b) { return a + b; }, 0);
        czvPairs.forEach(function (p, k) { parts.push([p, czvAmt * cw[k] / cws]); });
      } else {
        others = pairs;
      }
      var rest = cell - czvAmt;
      var ow = others.map(function () { return 0.2 + rnd(); }), ows = ow.reduce(function (a, b) { return a + b; }, 0);
      others.forEach(function (p, k) { parts.push([p, rest * ow[k] / ows]); });
      // округлення з точним підсумком клітинки (залишок — на найбільший рядок не-ЦЗВ)
      var s = 0, big = -1;
      parts.forEach(function (pt, k) { pt[1] = r2(pt[1]); s += pt[1]; if (pt[0][0] !== CZV && (big < 0 || pt[1] > parts[big][1])) big = k; });
      if (big >= 0) parts[big][1] = r2(parts[big][1] + cell - s);
      parts.forEach(function (pt) {
        var cl = pt[0][0], itm = pt[0][1];
        var sub = itm[3] ? GP + itm[3] : '';
        var path = g.path + sub;
        var ids = path.split(GP).map(gid).join(GP);
        var price = itm[2] * (0.97 + 0.06 * rnd());
        rows.push({
          month: month, organization: ORG, group_path: path, group_ids: ids,
          item_code: '00-' + gid(itm[0]).slice(0, 6).toUpperCase(), item: itm[0], unit: itm[1],
          counterparty_code: '00-' + gid(cl).slice(0, 7).toUpperCase(), counterparty: cl,
          qty: Math.round(pt[1] / price * 1000) / 1000, amount: pt[1]
        });
      });
    });
    var sum = r2(rows.reduce(function (a, r) { return a + r.amount; }, 0));
    var lastDay = new Date(Date.UTC(2026, m + 1, 0)).getUTCDate();
    payloads.push({
      token: '<секретний ключ>', dataset: 'sales', organization: ORG,
      period_from: month, period_to: '2026-' + ('0' + (m + 1)).slice(-2) + '-' + lastDay,
      generated_at: '2026-11-11T13:30:00+02:00', mode: 'replace_period',
      control: { rows: rows.length, amount: sum }, rows: rows
    });
  }
  return payloads;
}

/** Завантажує тестові дані в Таблицю тим самим шляхом, що й дані з 1С. */
function loadTestData() {
  var res = generateTestPayloads_().map(function (p) { return p.period_from + ': ' + JSON.stringify(ingestPayload_(p, 'test')); });
  console.log(res.join('\n'));
  // без діалогових вікон — можна запускати і з меню, і з редактора (результат у "Журналі виконання")
  try { getSpreadsheet_().toast('Тестові дані завантажено. Деталі — в аркуші _log.', 'ТРІБО', 10); } catch (e) {}
}

/**
 * Перевіряє наскрізний HTTP-шлях: надсилає листопадовий пакет на опублікований URL прийому.
 * Перед запуском: Властивості скрипта → INGEST_URL = URL веб-застосунку (…/exec).
 */
function testHttpPost() {
  var props = PropertiesService.getScriptProperties();
  var url = props.getProperty('INGEST_URL');
  if (!url) throw new Error('Додайте властивість INGEST_URL (URL веб-застосунку прийому)');
  var p = generateTestPayloads_()[10];
  p.token = props.getProperty('TOKEN');
  var resp = UrlFetchApp.fetch(url, { method: 'post', contentType: 'application/json; charset=utf-8',
    payload: JSON.stringify(p), muteHttpExceptions: true, followRedirects: true });
  console.log(resp.getResponseCode() + ' ' + resp.getContentText());
}

/** Звіряє дані в Таблиці з еталоном 9.5. */
function checkAgainstEtalon() {
  var ss = getSpreadsheet_();
  var ds = DATASETS.sales, sh = ss.getSheetByName(ds.sheet);
  var n = sh.getLastRow() - 1;
  if (n < 1) throw new Error('Аркуш sales порожній');
  var vals = sh.getRange(2, 1, n, ds.columns.length).getValues();
  var disp = sh.getRange(2, 1, n, 1).getDisplayValues();
  var gsh = ss.getSheetByName(CONFIG.SHEET_GROUPS), gmap = {};
  if (gsh.getLastRow() > 1) gsh.getRange(2, 1, gsh.getLastRow() - 1, 4).getValues().forEach(function (r) { gmap[r[0]] = r[3]; });
  var ci = function (c) { return ds.columns.indexOf(c); };
  var tot = 0, oct = 0, nov = 0, byG = {}, byC = {};
  vals.forEach(function (r, i) {
    var mo = disp[i][0];
    if (mo.slice(0, 4) !== String(ETALON.year)) return;
    var a = r[ci('amount')];
    tot += a;
    if (mo.slice(5, 7) === '10') oct += a;
    if (mo.slice(5, 7) === '11') nov += a;
    var g = gmap[r[ci('group_path')]] || r[ci('group_path')];
    byG[g] = (byG[g] || 0) + a;
    var c = r[ci('counterparty')];
    byC[c] = (byC[c] || 0) + a;
  });
  var t = function (x) { return Math.round(x / 100) / 10; };
  var lines = [];
  var chk = function (name, got, exp) { lines.push((Math.abs(got - exp) < 0.051 ? '✔ ' : '✘ ') + name + ': ' + got + ' (еталон ' + exp + ')'); };
  chk('Рік, тис. грн', t(tot), ETALON.yearTotal);
  chk('Жовтень', t(oct), ETALON.oct);
  chk('Листопад', t(nov), ETALON.nov);
  // група шукається за початком назви ("(2.10)" знайде і "(2.10) - Комплект…"); ігноруються дефіси й пробіли
  var norm = function (x) { return String(x).replace(/[\s\-–]+/g, ' ').trim().toLowerCase(); };
  Object.keys(ETALON.shares).forEach(function (g) {
    var key = norm(g).split(' ')[0] === 'мпс' ? 'мпс' : norm(g).split(' ')[0];
    var sum = 0;
    Object.keys(byG).forEach(function (n) { if (norm(n) === norm(g) || norm(n).split(' ')[0] === key) sum += byG[n]; });
    chk('Частка ' + g + ', %', Math.round(sum / tot * 1000) / 10, ETALON.shares[g]);
  });
  var top = Object.keys(byC).sort(function (a, b) { return byC[b] - byC[a]; })[0];
  lines.push((top === ETALON.topClient ? '✔ ' : '✘ ') + 'Найбільший клієнт: ' + top + ' — ' + Math.round(byC[top] / 1000) + ' тис. грн (еталон ' + ETALON.topClientAmount + ')');
  console.log(lines.join('\n'));
  try { getSpreadsheet_().toast(lines.join('  '), 'Звірка з еталоном', 30); } catch (e) {}
  return lines;
}

if (typeof module !== 'undefined') module.exports = { generateTestPayloads_: generateTestPayloads_, ETALON: ETALON };
