/**
 * ТРІБО · ЗАВАНТАЖУВАЧ. Вставляється в Apps Script (проєкт Таблиці «Трібо Дані») ОДИН РАЗ.
 *
 * Увесь робочий код береться з GitHub: https://github.com/TriboDashbords/tribo-dashboard
 * Зміни робляться там і підхоплюються автоматично (протягом ~2 хвилин) — цей файл міняти не потрібно.
 *
 * Функції нижче — лише "заглушки": Google викликає їх, а вони передають виклик у код з GitHub.
 */

var TRIBO_SRC = 'https://raw.githubusercontent.com/TriboDashbords/tribo-dashboard/main/';
var TRIBO_FRESH_SEC = 120; // як часто перевіряти GitHub на нову версію

function doGet(e) { return tribo_().doGet(e); }
function doPost(e) { return tribo_().doPost(e); }
function api(name, args) { return tribo_().api(name, args); }
function processInbox() { return tribo_().processInbox(); }
function setup() { return tribo_().setup(); }
function showToken() { return tribo_().showToken(); }
function run() { return tribo_().run(); }

function onOpen() {
  // Меню будується тут, без звернень до GitHub (простому тригеру onOpen це заборонено)
  SpreadsheetApp.getUi().createMenu('ТРІБО')
    .addItem('Налаштувати (setup)', 'setup')
    .addItem('Показати токен для 1С', 'showToken')
    .addItem('Обробити папку Inbox зараз', 'processInbox')
    .addItem('Виконати службову задачу (run)', 'run')
    .addToUi();
}

var TRIBO_ = null;
function tribo_() {
  if (!TRIBO_) TRIBO_ = new Function('triboLoad', triboLoad('dist/server.js') + '\n;return TRIBO_API;')(triboLoad);
  return TRIBO_;
}

/** Читає файл з GitHub. Кеш: свіжа копія — 2 хв; запасна — 6 год (якщо GitHub недоступний). */
function triboLoad(path) {
  var cache = CacheService.getScriptCache();
  var fresh = triboCacheGet_(cache, 'f:' + path);
  if (fresh !== null) return fresh;
  var text = null;
  try {
    var r = UrlFetchApp.fetch(TRIBO_SRC + path + '?t=' + Math.floor(Date.now() / 60000), { muteHttpExceptions: true });
    if (r.getResponseCode() === 200) text = r.getContentText('UTF-8');
  } catch (e) {}
  if (text !== null) {
    triboCachePut_(cache, 'f:' + path, text, TRIBO_FRESH_SEC);
    triboCachePut_(cache, 'l:' + path, text, 21600);
    return text;
  }
  var last = triboCacheGet_(cache, 'l:' + path);
  if (last !== null) return last;
  throw new Error('Не вдалося завантажити код з GitHub: ' + path);
}

function triboCachePut_(cache, key, s, ttl) {
  var size = 90000, n = Math.ceil(s.length / size) || 1, items = {};
  for (var i = 0; i < n; i++) items[key + '#' + i] = s.slice(i * size, (i + 1) * size);
  items[key] = String(n);
  cache.putAll(items, ttl);
}

function triboCacheGet_(cache, key) {
  var n = cache.get(key);
  if (!n) return null;
  var keys = [];
  for (var i = 0; i < parseInt(n, 10); i++) keys.push(key + '#' + i);
  var parts = cache.getAll(keys), s = '';
  for (var j = 0; j < keys.length; j++) { if (parts[keys[j]] === undefined) return null; s += parts[keys[j]]; }
  return s;
}
