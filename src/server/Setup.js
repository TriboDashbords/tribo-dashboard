/**
 * Setup.gs — первинне налаштування, меню, архівація.
 * Запустіть setup() один раз після вставки коду (і повторно — після додавання нових наборів у Config.gs).
 */

function onOpen() {
  // У меню — лише функції-"заглушки" із завантажувача (loader/Code.gs)
  SpreadsheetApp.getUi().createMenu('ТРІБО')
    .addItem('Налаштувати (setup)', 'setup')
    .addItem('Показати токен для 1С', 'showToken')
    .addItem('Обробити папку Inbox зараз', 'processInbox')
    .addItem('Виконати службову задачу (run)', 'run')
    .addToUi();
}

function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var props = PropertiesService.getScriptProperties();
  props.setProperty('SPREADSHEET_ID', ss.getId());
  if (!props.getProperty('TOKEN')) props.setProperty('TOKEN', newToken_());
  ss.setSpreadsheetTimeZone('Europe/Kiev');

  // службові аркуші та аркуші наборів
  getOrCreateSheet_(ss, CONFIG.SHEET_LOG,
    ['time', 'source', 'dataset', 'organization', 'period', 'mode', 'rows', 'amount', 'status', 'message', 'duration_ms']);
  getOrCreateSheet_(ss, CONFIG.SHEET_GROUPS,
    ['group_path', 'group_ids', 'product_group', 'display_name', 'status', 'first_seen']);
  getOrCreateSheet_(ss, CONFIG.SHEET_CONTROL,
    ['dataset', 'organization', 'period_from', 'period_to', 'rows', 'amount', 'source', 'received_at']);
  Object.keys(DATASETS).forEach(function (name) {
    getOrCreateSheet_(ss, DATASETS[name].sheet, DATASETS[name].columns);
    updateDatasetRegistry_(ss, name, DATASETS[name], null);
  });
  // користувачі дашборду: власник додається автоматично
  var us = getOrCreateSheet_(ss, AUTH.USERS_SHEET, ['email', 'name', 'active', 'added', 'password']);
  var owner = '';
  try { owner = Session.getEffectiveUser().getEmail(); } catch (e) {}
  if (owner && us.getLastRow() < 2) us.appendRow([owner, 'Власник', 'так', nowIso_()]);
  var def = ss.getSheetByName('Аркуш1') || ss.getSheetByName('Sheet1') || ss.getSheetByName('Лист1');
  if (def && ss.getSheets().length > 1 && def.getLastRow() === 0) ss.deleteSheet(def);

  // папки Drive для резервного шляху
  var inbox = getOrCreateFolder_(DriveApp.getRootFolder(), CONFIG.INBOX_FOLDER_NAME);
  getOrCreateFolder_(inbox, CONFIG.INBOX_DONE_SUBFOLDER);
  getOrCreateFolder_(inbox, CONFIG.INBOX_ERROR_SUBFOLDER);
  props.setProperty('INBOX_FOLDER_ID', inbox.getId());

  // тригер: перевіряти Inbox кожні 15 хв
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'processInbox') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('processInbox').timeBased().everyMinutes(15).create();

  log_(ss, { source: 'setup', status: 'OK', message: 'Налаштування виконано. ID Таблиці: ' + ss.getId() });
  // Без діалогових вікон: при запуску з редактора вікно відкривалось би у вкладці Таблиці
  // і скрипт чекав би натискання "OK" до тайм-ауту.
  ss.toast('Налаштування виконано. ID Таблиці — в аркуші _log.', 'ТРІБО', 10);
  console.log('Готово. ID Таблиці (для проєкту дашборду): ' + ss.getId());
}

function showToken() {
  var t = PropertiesService.getScriptProperties().getProperty('TOKEN') || '(не задано — запустіть setup)';
  console.log('Токен для 1С: ' + t);
  try { getSpreadsheet_().toast('Токен — у "Журналі виконання" редактора', 'ТРІБО', 10); } catch (e) {}
  return t;
}

function rotateToken() {
  PropertiesService.getScriptProperties().setProperty('TOKEN', newToken_());
  log_(getSpreadsheet_(), { source: 'setup', status: 'OK', message: 'Токен змінено. Передайте новий програмісту 1С.' });
  showToken();
}

function newToken_() {
  return Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
}

function getOrCreateFolder_(parent, name) {
  var it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

/**
 * Видаляє тестові дані: рядки продажів, контрольні суми і продуктові групи (_groups).
 * Запускати з редактора (функція clearTestData) перед завантаженням справжніх даних.
 * Журнал _log не чіпає.
 */
function clearTestData() {
  var ss = getSpreadsheet_();
  [DATASETS.sales.sheet, CONFIG.SHEET_CONTROL, CONFIG.SHEET_GROUPS].forEach(function (n) {
    var sh = ss.getSheetByName(n);
    if (sh && sh.getLastRow() > 1) safeDeleteRows_(sh, 2, sh.getLastRow() - 1);
  });
  updateDatasetRegistry_(ss, 'sales', DATASETS.sales, { updated: nowIso_(), generated_at: '' });
  log_(ss, { source: 'manual', dataset: 'sales', status: 'OK', message: 'Тестові дані видалено (sales, _control, _groups)' });
  console.log('Тестові дані видалено.');
}

/** Стара назва — залишена для меню. */
function clearSalesData() { clearTestData(); }

/**
 * Архівація: переносить рядки року в окрему Таблицю "ТРІБО архів <рік>" (на тому ж Диску)
 * і видаляє їх з основної. Робити, коли в _log з'являється попередження про кількість клітинок.
 */
function archiveYearPrompt() {
  var ui = SpreadsheetApp.getUi();
  var r = ui.prompt('Архівація', 'Який рік винести в архів? (напр. 2026)', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  var year = String(r.getResponseText()).trim();
  if (!/^\d{4}$/.test(year)) { ui.alert('Потрібен рік з 4 цифр'); return; }
  var url = archiveYear(year);
  ui.alert('Готово', 'Архів: ' + url, ui.ButtonSet.OK);
}

function archiveYear(year) {
  var ss = getSpreadsheet_();
  var arch = SpreadsheetApp.create('ТРІБО архів ' + year);
  Object.keys(DATASETS).forEach(function (name) {
    var ds = DATASETS[name];
    var sh = ss.getSheetByName(ds.sheet);
    if (!sh || sh.getLastRow() < 2) return;
    var pCol = ds.columns.indexOf(ds.periodField);
    var all = sh.getRange(2, 1, sh.getLastRow() - 1, ds.columns.length).getValues();
    var disp = sh.getRange(2, pCol + 1, sh.getLastRow() - 1, 1).getDisplayValues();
    var take = [], keep = [];
    all.forEach(function (row, i) { (String(disp[i][0]).slice(0, 4) === year ? take : keep).push(row); });
    if (!take.length) return;
    var ash = getOrCreateSheet_(arch, ds.sheet, ds.columns);
    writeRows_(ash, ds, 2, take);
    sh.getRange(2, 1, all.length, ds.columns.length).clearContent();
    if (keep.length) writeRows_(sh, ds, 2, keep);
    if (sh.getMaxRows() > keep.length + 2) sh.deleteRows(keep.length + 3, sh.getMaxRows() - keep.length - 2);
  });
  log_(ss, { source: 'archive', status: 'OK', message: 'Рік ' + year + ' винесено: ' + arch.getUrl() });
  return arch.getUrl();
}
