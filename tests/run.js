// Наскрізний тест у Node з імітацією сервісів Google Apps Script.
// Запуск: node build.js && node tests/run.js
const fs = require('fs'), crypto = require('crypto'), path = require('path');
const ROOT = path.join(__dirname, '..');
let failed = 0;
const ok = (cond, msg) => { console.log((cond ? '  ✔ ' : '  ✘ ') + msg); if (!cond) failed++; };

// ---------- імітація Google ----------
function fmtCell(v) { return v === null || v === undefined ? '' : String(v); }
class Range {
  constructor(sh, r, c, nr, nc) { Object.assign(this, { sh, r, c, nr: nr || 1, nc: nc || 1 }); }
  _chk() { if (this.r + this.nr - 1 > this.sh.maxRows) throw new Error('out of bounds rows'); if (this.c + this.nc - 1 > this.sh.maxCols) throw new Error('out of bounds cols'); }
  getValues() { const o = []; for (let i = 0; i < this.nr; i++) { const row = []; for (let j = 0; j < this.nc; j++) { const v = (this.sh.data[this.r - 1 + i] || [])[this.c - 1 + j]; row.push(v === undefined ? '' : v); } o.push(row); } return o; }
  getDisplayValues() { return this.getValues().map(r => r.map(fmtCell)); }
  getValue() { return this.getValues()[0][0]; }
  setValue(v) { return this.setValues([[v]]); }
  setValues(v) { this._chk(); for (let i = 0; i < this.nr; i++) { const rr = this.sh.data[this.r - 1 + i] || (this.sh.data[this.r - 1 + i] = []); for (let j = 0; j < this.nc; j++) rr[this.c - 1 + j] = v[i][j]; } return this; }
  setNumberFormats() { return this; } setNumberFormat() { return this; } setFontWeight() { return this; }
  clearContent() { for (let i = 0; i < this.nr; i++) { const rr = this.sh.data[this.r - 1 + i]; if (rr) for (let j = 0; j < this.nc; j++) rr[this.c - 1 + j] = ''; } return this; }
}
class Sheet {
  constructor(name) { this.name = name; this.data = []; this.maxRows = 1000; this.maxCols = 26; this.frozen = 0; }
  getName() { return this.name; }
  getRange(r, c, nr, nc) { return new Range(this, r, c, nr, nc); }
  getLastRow() { for (let i = this.data.length - 1; i >= 0; i--) if (this.data[i] && this.data[i].some(v => v !== '' && v !== undefined)) return i + 1; return 0; }
  getLastColumn() { return Math.max(0, ...this.data.map(r => r ? r.length : 0)); }
  getMaxRows() { return this.maxRows; } getMaxColumns() { return this.maxCols; }
  setFrozenRows(n) { this.frozen = n; } getFrozenRows() { return this.frozen; }
  deleteColumns(c, n) { this.maxCols -= n; }
  insertColumnsAfter(c, n) { this.maxCols += n; }
  deleteRows(r, n) { if (this.maxRows - n <= this.frozen) throw new Error('Не можна видалити всі незакріплені рядки.'); this.data.splice(r - 1, n); this.maxRows -= n; }
  deleteRow(r) { this.deleteRows(r, 1); }
  insertRowsAfter(r, n) { this.maxRows += n; }
  appendRow(row) { const lr = this.getLastRow(); if (lr + 1 > this.maxRows) this.maxRows = lr + 1; this.data[lr] = row.slice(); }
}
const SS = { sheets: [], getSheetByName(n) { return this.sheets.find(s => s.name === n) || null; },
  insertSheet(n) { const s = new Sheet(n); this.sheets.push(s); return s; }, getSheets() { return this.sheets; },
  getId() { return 'TEST_SS'; }, toast() {}, setSpreadsheetTimeZone() {}, deleteSheet(s) { this.sheets = this.sheets.filter(x => x !== s); } };
const store = { cache: new Map(), props: new Map(), mails: [] };
const G = {
  console,
  SpreadsheetApp: { openById: () => SS, getActiveSpreadsheet: () => SS, getUi() { throw new Error('no ui'); } },
  PropertiesService: { getScriptProperties: () => ({
    getProperty: k => store.props.has(k) ? store.props.get(k) : null, setProperty: (k, v) => store.props.set(k, String(v)),
    deleteProperty: k => store.props.delete(k), getProperties: () => Object.fromEntries(store.props) }) },
  CacheService: { getScriptCache: () => ({
    get: k => store.cache.has(k) ? store.cache.get(k) : null, put: (k, v) => store.cache.set(k, String(v)), remove: k => store.cache.delete(k),
    putAll: o => Object.entries(o).forEach(([k, v]) => store.cache.set(k, String(v))),
    getAll: ks => Object.fromEntries(ks.filter(k => store.cache.has(k)).map(k => [k, store.cache.get(k)])) }) },
  LockService: { getScriptLock: () => ({ tryLock: () => true, releaseLock() {} }) },
  MailApp: { sendEmail: m => store.mails.push(m) },
  Session: { getScriptTimeZone: () => 'Europe/Kiev', getEffectiveUser: () => ({ getEmail: () => 'owner@example.com' }) },
  Utilities: {
    DigestAlgorithm: { SHA_256: 'sha256', MD5: 'md5' }, Charset: { UTF_8: 'utf8' },
    computeDigest: (alg, s) => Array.from(crypto.createHash(alg).update(String(s), 'utf8').digest()),
    base64EncodeWebSafe: b => Buffer.from(b.map(x => x & 255)).toString('base64').replace(/\+/g, '-').replace(/\//g, '_'),
    getUuid: () => crypto.randomUUID(),
    formatDate: (d, tz, f) => f === 'yyyy' ? String(d.getFullYear()) : f === 'M' ? String(d.getMonth() + 1) : f === 'yyyy-MM-dd' ? d.toISOString().slice(0, 10) : d.toISOString()
  },
  ContentService: { MimeType: { JSON: 'json' }, createTextOutput: s => ({ setMimeType() { return { text: s }; } }) },
  HtmlService: { createHtmlOutput: h => ({ setTitle() { return this; }, addMetaTag() { return this; }, html: h }) },
  DriveApp: {}, ScriptApp: {},
  GmailApp: (function () {
    const labels = {}; const mk = n => (labels[n] = labels[n] || { n, threads: [] });
    const att = { getName: () => 'звіт.xlsx', copyBlob: () => ({}) };
    const msg = { getFrom: () => 'Stranger <stranger@example.com>', getSubject: () => 'Звіт', getAttachments: () => [att] };
    const th = { labels: [], getMessages: () => [msg], addLabel(l) { this.labels.push(l.n); }, markRead() {} };
    return { _th: th, getUserLabelByName: n => labels[n] || null, createLabel: mk, search: () => [th] };
  })()
};

// ---------- завантаження зібраного коду так само, як у loader/Code.gs ----------
const triboLoad = p => fs.readFileSync(path.join(ROOT, p), 'utf8');
const src = triboLoad('dist/server.js');
const names = Object.keys(G);
const T = new Function(...names, 'triboLoad', src + '\n;return {TRIBO_API, generateTestPayloads_, ingestPayload_, validatePayload_, readXlsxTable_, parse1cSalesTable_, findKeyCaseInsensitive_, syncGroups_, getOrCreateSheet_, safeDeleteRows_, parsePlanFactTable_, processGmail_};')(...names.map(n => G[n]), triboLoad);

console.log('1. Завантаження коду');
['doGet', 'doPost', 'api', 'processInbox', 'setup', 'onOpen', 'showToken', 'run'].forEach(f => ok(typeof T.TRIBO_API[f] === 'function', 'TRIBO_API.' + f));

console.log('2. Прийом тестових даних');
store.props.set('SPREADSHEET_ID', 'TEST_SS'); store.props.set('TOKEN', 'secret');
const users = T.getOrCreateSheet_(SS, '_users', ['email', 'name', 'active', 'added', 'password']);
users.appendRow(['Owner@Example.com', 'Власник', 'так', '', 'Pa55word']);
users.appendRow(['blocked@example.com', 'X', 'ні', '']);
const P = T.generateTestPayloads_();
P.forEach(p => { const r = T.ingestPayload_(p, 'test'); if (!r.ok) console.log(r); });
ok(SS.getSheetByName('sales').getLastRow() - 1 === P.reduce((s, p) => s + p.rows.length, 0), 'усі рядки записано');
// повторне надсилання не створює дублів
T.ingestPayload_(P[10], 'test');
ok(SS.getSheetByName('sales').getLastRow() - 1 === P.reduce((s, p) => s + p.rows.length, 0), 'replace_period без дублів');
// doPost з невірним токеном
const bad = T.TRIBO_API.doPost({ postData: { contents: JSON.stringify(Object.assign({}, P[0], { token: 'x' })) } });
ok(/Невірний токен/.test(bad.text), 'doPost відхиляє невірний токен');
const good = T.TRIBO_API.doPost({ postData: { contents: JSON.stringify(Object.assign({}, P[0], { token: 'secret' })) } });
ok(JSON.parse(good.text).ok === true, 'doPost приймає правильний токен');

console.log('3. Безпечне видалення рядків');
const tmp = T.getOrCreateSheet_(SS, '_tmp', ['a']); tmp.appendRow([1]); tmp.appendRow([2]);
let thrown = null; try { T.safeDeleteRows_(tmp, 2, tmp.getLastRow() - 1); } catch (e) { thrown = e; }
ok(!thrown, 'видалення всіх рядків даних не падає');

console.log('4. Вхід за кодом');
const api = T.TRIBO_API.api;
let e1 = null; try { api('init', ['']); } catch (e) { e1 = e.message; }
ok(e1 === 'AUTH_REQUIRED', 'без токена — AUTH_REQUIRED');
api('requestCode', ['blocked@example.com']);
ok(store.mails.length === 0, 'недозволеній адресі код не надсилається');
api('requestCode', ['someone@example.com']);
ok(store.mails.length === 0, 'невідомій адресі код не надсилається');
api('requestCode', [' owner@example.com ']);
ok(store.mails.length === 1, 'дозволеній адресі код надіслано');
const code = /(\d{6})/.exec(store.mails[0].subject)[1];
let e2 = null; try { api('verifyCode', ['owner@example.com', '000000' === code ? '111111' : '000000']); } catch (e) { e2 = e.message; }
ok(e2 === 'Невірний код', 'невірний код відхиляється');
const sess = api('verifyCode', ['owner@example.com', code]);
ok(sess.token && sess.token.length >= 32, 'видано токен сесії');
let e3 = null; try { api('verifyCode', ['owner@example.com', code]); } catch (e) { e3 = e.message; }
ok(!!e3, 'код одноразовий');

console.log('4b. Вхід за паролем');
let e5 = null; try { api('login', ['owner@example.com', 'wrong']); } catch (e) { e5 = e.message; }
ok(e5 === 'Невірна пошта або пароль', 'невірний пароль відхиляється');
let e6 = null; try { api('login', ['blocked@example.com', '']); } catch (e) { e6 = e.message; }
ok(e6 === 'Невірна пошта або пароль', 'заблокований / без пароля — відмова');
const s2 = api('login', [' OWNER@example.com', 'Pa55word']);
ok(api('init', [s2.token]).yearPeriod.total > 0, 'вхід за паролем дає доступ до даних');

console.log('5. Дані дашборду');
const init = api('init', [sess.token]);
const yr = init.yearPeriod;
ok(Math.abs(yr.total / 1000 - 104500.0) < 0.05, 'рік = 104 500,0 тис (тест): ' + (yr.total / 1000).toFixed(1));
ok(Math.abs(yr.groupsSum - yr.total) < 1 && Math.abs(yr.clientsSum - yr.total) < 1, 'групи = клієнти = підсумок');
const mps = yr.groups.find(g => g.name === 'МПС');
ok(mps && Math.abs(mps.amount / yr.total * 100 - 45) < 0.06, 'частка МПС 45% (тест)');
ok(yr.clients[0].name === 'ТОВ "Ключовий клієнт (тест)"' && Math.round(yr.clients[0].amount / 1000) === 25000, 'найбільший клієнт (тест)');
const items = api('clientItems', [sess.token, init.year, 0, yr.clients[0].key]);
ok(items.length > 0 && items[0].amount >= items[items.length - 1].amount, 'продукти клієнта відсортовані');
const p10 = api('period', [sess.token, init.year, 10]);
ok(Math.abs(p10.total / 1000 - 12000) < 0.05 && p10.control !== null, 'жовтень 12 000,0 і контрольна сума є');

console.log('6. Вихід і блокування');
api('logout', [sess.token]);
let e4 = null; try { api('init', [sess.token]); } catch (e) { e4 = e.message; }
ok(e4 === 'AUTH_REQUIRED', 'після виходу токен недійсний');

console.log('6b. Старий аркуш _users з 4 стовпцями');
SS.sheets = SS.sheets.filter(s => s.name !== '_users');
const old = T.getOrCreateSheet_(SS, '_users', ['email', 'name', 'active', 'added']);
ok(old.getMaxColumns() === 4, 'імітація: 4 стовпці');
T.TRIBO_API.doGet({ parameter: {} });
ok(old.getMaxColumns() >= 5 && old.getRange(1, 5).getValue() === 'password', 'стовпець password додано автоматично');
ok(old.getRange(2, 1).getValue() === 'owner@example.com', 'власника додано автоматично');
old.getRange(2, 5).setValue('TestPass1');
ok(!!api('login', ['owner@example.com', 'TestPass1']).token, 'вхід із паролем після виправлення');

console.log('6c. План-факт (вигадані дані)');
{
  const H = [null, 'Валюта', null, null, null, 'Количество (план)', 'Количество (факт)', 'Сумма (план продажи)', 'Сумма (факт продажи)', 'Сумма (факт продаж с НДС)', 'Сумма (приход проплат)', 'Сумма (долг)', 'Цена себестоимости (маржинальная)', 'Сумма себестоимости (план)', 'Сумма себестоимости (факт)', 'Прибыль (план)', 'Прибыль (факт)', 'Рентабельность (план)', 'Рентабельность (факт)'];
  const row = (level, cells) => ({ level, cells });
  const it = (lvl, code, name, qp, qf, sp, sf, cp, cf) => row(lvl, [null, code, name, null, null, qp, qf, sp, sf, sf * 1.2, null, null, null, cp, cf, sp - cp, sf - cf]);
  const tbl = [
    row(0, [null, 'Параметры :', 'Период с:', '01.10.2026']), row(0, [null, null, 'по:', '31.10.2026']),
    row(0, [null, null, 'Организация :', 'ORG']), row(0, H), row(0, [null, 'Клонтрагент']), row(0, [null, 'Код', 'Номенклатура']),
    row(0, [null, 'Регіон А']), row(1, [null, 'Клієнт 1,', null, null, null, 0, 0, 0, 0, 0, 500, 100]),
    row(2, [null, 'Група X']), row(3, [null, 'Підгрупа Y']), it(4, 101, 'Товар 1', 1, 1, 1000, 800, 500, 400), it(4, 102, 'Товар 2', 2, 1, 2000, 900, 1000, 600),
    row(2, [null, 'Група Z']), it(3, 103, 'Товар 3', 1, 0, 300, 0, 100, 0),
    row(1, [null, 'Клієнт 2', null, null, null, 0, 0, 0, 0, 0, 0, -50]), row(2, [null, 'Група X']), row(3, [null, 'Підгрупа Y']), it(4, 101, 'Товар 1', 1, 2, 1000, 1600, 500, 800),
    row(0, [null, 'Общий итог', null, null, null, 0, 0, 4300, 3300])
  ];
  const r = T.parsePlanFactTable_(tbl, { organization: 'ORG' });
  ok(r.errors.length === 0 && r.payloads[0].rows.length === 4 && r.payloads[1].rows.length === 2, 'план-факт розібрано: 4 товари, 2 клієнти');
  ok(r.payloads[1].rows[0].counterparty === 'Клієнт 1', 'кома в кінці назви клієнта прибрана');
  r.payloads.forEach(p => { const x = T.ingestPayload_(p, 'test'); ok(x.ok, 'запис ' + p.dataset + (x.ok ? '' : ': ' + x.error)); });
  ok(JSON.stringify(api('init', [s2.token]).planfactMonths) === '["2026-10"]', 'місяці план-факту в init');
  const v = api('planfact', [s2.token, '2026-10']);
  ok(v.totals.f === 3300 && v.totals.pr === 3300 - 1800 && v.control === 3300, 'підсумки факту, маржа і контроль');
  ok(v.groups.length === 1 && v.groups[0].f === 3300 && v.groups[0].it[0].f === 2400 && v.groups[0].it[0].q === 3, 'групи: товари без плану відкинуто, ранжування');
  const c1 = v.clients.find(c => c.n === 'Клієнт 2'), c0 = v.clients.find(c => c.n === 'Клієнт 1');
  ok(c1 && c1.debt === -50 && c1.g[0].it[0].pr === 800 && c0.debt === 100, 'клієнт: борг, групи і продукти');
  ok(v.clients[0].f >= v.clients[1].f, 'клієнти від більшого до меншого');
  // повторне завантаження з іншим написанням організації не дублює рядки
  const p2 = JSON.parse(JSON.stringify(r.payloads[0])); p2.organization = '"ORG"'; p2.rows.forEach(x => x.organization = '"ORG"');
  ok(T.ingestPayload_(p2, 'test').ok && SS.getSheetByName('planfact').getLastRow() - 1 === 4, 'інше написання організації — заміна без дублів');
}

console.log('6d. Пошта: листи від невідомих відправників не обробляються');
{
  const n = T.processGmail_(Date.now());
  const last = SS.getSheetByName('_log').data.filter(r => r && r[1] === 'gmail').pop();
  ok(n === 0 && G.GmailApp._th.labels[0] === 'tribo-error' && /_senders/.test(last[9]), 'чужий лист пропущено і позначено tribo-error');
}

console.log('7. Сторінка');
const page = T.TRIBO_API.doGet({ parameter: {} });
ok(/Вхід у дашборд/.test(page.html) && /Пароль/.test(page.html) && /google\.script\.run/.test(page.html), 'doGet віддає сторінку зі входом');
ok(/"ok":true/.test(T.TRIBO_API.doGet({ parameter: { health: '1' } }).text), '?health=1 відповідає');

console.log('8. Розбір Excel 1С (синтетичний файл, якщо є tests/fixtures)');
const fx = path.join(__dirname, 'fixtures', 'report_1c');
if (fs.existsSync(fx)) {
  const names2 = []; (function w(d) { for (const f of fs.readdirSync(d)) { const p = path.join(d, f); fs.statSync(p).isDirectory() ? w(p) : names2.push(path.relative(fx, p)); } })(fx);
  const map = {}; names2.forEach(n => map[n] = n);
  const table = T.readXlsxTable_(p => { const k = T.findKeyCaseInsensitive_(map, p); return k ? fs.readFileSync(path.join(fx, k), 'utf8') : null; }, names2);
  const res = T.parse1cSalesTable_(table, { organization: 'ORG', fallbackYear: 2026 });
  ok(res.errors.length === 0 && res.payloads.length > 0, 'Excel розібрано: ' + res.payloads.map(p => p.period_from + ':' + p.rows.length).join(', '));
  res.payloads.forEach(p => { const v = T.validatePayload_(p); ok(v.errors.length === 0, 'контрольна сума ' + p.period_from); });
} else console.log('  (пропущено)');

console.log(failed ? `\nПОМИЛОК: ${failed}` : '\nУсі перевірки пройдено.');
process.exit(failed ? 1 : 0);
