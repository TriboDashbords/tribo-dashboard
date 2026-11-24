/**
 * Api.js — точки входу. Завантажувач (loader/Code.gs) передає сюди виклики:
 *   doGet  — сторінка дашборду (або ?health=1 — перевірка);
 *   doPost — прийом даних з 1С (Ingest.js);
 *   api    — запити сторінки дашборду (вхід і дані);
 *   run    — "службова задача": що саме робить, визначається тут, у RUN_TASK.
 * `triboLoad(path)` — функція завантажувача, що читає файл із GitHub (з кешем).
 */

function doGet(e) {
  if (e && e.parameter && e.parameter.health) return json_({ ok: true, service: 'tribo', time: nowIso_() });
  try { ensureUsersSheet_(); } catch (err) { console.error('ensureUsersSheet_: ' + err); }
  return HtmlService.createHtmlOutput(triboLoad('dist/index.html'))
    .setTitle('ТРІБО · Продажі')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

var PUBLIC_API = {
  login: function (email, password) { return authPassword_(email, password); },
  requestCode: function (email) { return authRequestCode_(email); },
  verifyCode: function (email, code) { return authVerify_(email, code); },
  logout: function (token) { return authLogout_(token); }
};

// Перший аргумент кожного виклику — токен сесії
var PRIVATE_API = {
  init: function () { return apiInit(); },
  period: function (year, period) { return apiPeriod(year, period); },
  clientItems: function (year, period, clientKey) { return apiClientItems(year, period, clientKey); }
};

function api(name, args) {
  args = args || [];
  if (PUBLIC_API.hasOwnProperty(name)) return PUBLIC_API[name].apply(null, args);
  if (PRIVATE_API.hasOwnProperty(name)) {
    authCheck_(args[0]);
    return PRIVATE_API[name].apply(null, args.slice(1));
  }
  throw new Error('Невідома дія: ' + name);
}

/**
 * Службова задача, яку власник запускає кнопкою "run" у редакторі.
 * Claude змінює її вміст під конкретну потребу (перебудувати групи, звірка тощо).
 */
var RUN_TASK = function () { return statusReport_(); };

function statusReport_() {
  var ss = getSpreadsheet_(), lines = [];
  Object.keys(DATASETS).forEach(function (n) {
    var sh = ss.getSheetByName(DATASETS[n].sheet);
    lines.push(n + ': ' + (sh ? Math.max(0, sh.getLastRow() - 1) : 0) + ' рядків');
  });
  lines.push('користувачів: ' + Object.keys(allowedUsers_()).length);
  lines.push('клітинок у Таблиці: ' + countCells_(ss));
  console.log(lines.join('\n'));
  return lines;
}

var TRIBO_API = {
  doGet: doGet,
  doPost: doPost,
  api: api,
  processInbox: processInbox,
  setup: setup,
  onOpen: onOpen,
  showToken: showToken,
  run: function () { return RUN_TASK(); }
};
