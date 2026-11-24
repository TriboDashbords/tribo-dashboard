/**
 * Storage.gs — робота з Google Таблицею: аркуші, журнал, реєстр наборів, заміна періоду.
 */

function getSpreadsheet_() {
  var id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
}

function nowIso_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ssXXX");
}

/** Повертає аркуш; якщо його немає — створює із заголовками й прибирає зайві колонки. */
function getOrCreateSheet_(ss, name, headers) {
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
    var extraCols = sh.getMaxColumns() - headers.length;
    if (extraCols > 0) sh.deleteColumns(headers.length + 1, extraCols);
    var extraRows = sh.getMaxRows() - 2;
    if (extraRows > 0) sh.deleteRows(3, extraRows);
  }
  return sh;
}

/**
 * Безпечне видалення рядків: Google не дозволяє видалити всі незакріплені рядки аркуша,
 * тому, якщо після видалення їх не лишилося б, спершу додаємо один порожній рядок у кінці.
 */
function safeDeleteRows_(sh, row, count) {
  if (count <= 0) return;
  if (sh.getMaxRows() - count <= sh.getFrozenRows()) sh.insertRowsAfter(sh.getMaxRows(), 1);
  sh.deleteRows(row, count);
}

/** Рядок у журнал _log. */
function log_(ss, e) {
  try {
    var sh = getOrCreateSheet_(ss, CONFIG.SHEET_LOG,
      ['time', 'source', 'dataset', 'organization', 'period', 'mode', 'rows', 'amount', 'status', 'message', 'duration_ms']);
    sh.appendRow([nowIso_(), e.source || '', e.dataset || '', e.organization || '', e.period || '',
      e.mode || '', e.rows === undefined ? '' : e.rows, e.amount === undefined ? '' : e.amount,
      e.status || '', String(e.message || '').slice(0, 1000), e.duration_ms || '']);
    var over = sh.getLastRow() - 1 - CONFIG.LOG_MAX_ROWS;
    if (over > 0) sh.deleteRows(2, over);
  } catch (err) {
    console.error('log_ failed: ' + err);
  }
}

/** Оновлює рядок набору в реєстрі _datasets. */
function updateDatasetRegistry_(ss, name, ds, extra) {
  var sh = getOrCreateSheet_(ss, CONFIG.SHEET_DATASETS,
    ['dataset', 'sheet', 'columns', 'description', 'last_updated', 'generated_at', 'rows_total']);
  var last = sh.getLastRow();
  var rowIdx = -1;
  if (last > 1) {
    var names = sh.getRange(2, 1, last - 1, 1).getValues();
    for (var i = 0; i < names.length; i++) if (names[i][0] === name) { rowIdx = i + 2; break; }
  }
  var dataSheet = ss.getSheetByName(ds.sheet);
  var total = dataSheet ? Math.max(0, dataSheet.getLastRow() - 1) : 0;
  var row = [name, ds.sheet, ds.columns.join(', '), ds.description,
    extra && extra.updated ? extra.updated : '', extra && extra.generated_at ? extra.generated_at : '', total];
  if (rowIdx < 0) {
    sh.appendRow(row);
  } else {
    if (!extra || !extra.updated) { // не затирати дату оновлення при setup()
      var old = sh.getRange(rowIdx, 5, 1, 2).getValues()[0];
      row[4] = old[0]; row[5] = old[1];
    }
    sh.getRange(rowIdx, 1, 1, row.length).setValues([row]);
  }
}

/** Контрольні суми з 1С ("Итог"), по одному рядку на набір/організацію/період. */
function saveControl_(ss, c) {
  var sh = getOrCreateSheet_(ss, CONFIG.SHEET_CONTROL,
    ['dataset', 'organization', 'period_from', 'period_to', 'rows', 'amount', 'source', 'received_at']);
  var last = sh.getLastRow();
  if (last > 1) {
    var vals = sh.getRange(2, 1, last - 1, 4).getDisplayValues();
    for (var i = vals.length - 1; i >= 0; i--) {
      var v = vals[i];
      var overlap = v[0] === c.dataset && v[1] === c.organization &&
        !(v[3] < c.period_from || v[2] > c.period_to);
      if (overlap) safeDeleteRows_(sh, i + 2, 1);
    }
  }
  var target = sh.getLastRow() + 1;
  if (target > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), target - sh.getMaxRows());
  var r = sh.getRange(target, 1, 1, 8);
  r.setNumberFormats([['@', '@', '@', '@', '0', '0.00', '@', '@']]);
  r.setValues([[c.dataset, c.organization, c.period_from, c.period_to,
    c.rows === null ? '' : c.rows, c.amount === null ? '' : c.amount, c.source, nowIso_()]]);
}

/**
 * Замінює дані набору за період і організацію (mode = replace_period).
 * rows — масив масивів у порядку ds.columns. Повертає кількість видалених рядків.
 */
function replacePeriod_(ss, ds, organization, periodFrom, periodTo, rows) {
  var headers = ds.columns;
  var sh = getOrCreateSheet_(ss, ds.sheet, headers);
  var pCol = headers.indexOf(ds.periodField) + 1;
  var oCol = headers.indexOf(ds.orgField) + 1;
  var fromM = periodFrom.slice(0, 7), toM = periodTo.slice(0, 7);
  var last = sh.getLastRow();
  var deleted = 0;

  if (last > 1) {
    var pVals = sh.getRange(2, pCol, last - 1, 1).getDisplayValues();
    var oVals = sh.getRange(2, oCol, last - 1, 1).getDisplayValues();
    // знайти суцільні діапазони рядків, що підлягають заміні
    var runs = [], start = -1;
    for (var i = 0; i < pVals.length; i++) {
      var m = String(pVals[i][0]).slice(0, 7);
      var hit = m >= fromM && m <= toM && oVals[i][0] === organization;
      if (hit && start < 0) start = i;
      if (!hit && start >= 0) { runs.push([start, i - 1]); start = -1; }
      if (hit) deleted++;
    }
    if (start >= 0) runs.push([start, pVals.length - 1]);

    if (runs.length > 30) {
      // дані розкидані — швидше переписати аркуш повністю
      var all = sh.getRange(2, 1, last - 1, headers.length).getValues();
      var keep = all.filter(function (r, idx) {
        var m2 = String(pVals[idx][0]).slice(0, 7);
        return !(m2 >= fromM && m2 <= toM && oVals[idx][0] === organization);
      });
      sh.getRange(2, 1, last - 1, headers.length).clearContent();
      if (keep.length) writeRows_(sh, ds, 2, keep);
      if (sh.getMaxRows() > keep.length + 2) sh.deleteRows(keep.length + 3, sh.getMaxRows() - keep.length - 2);
    } else {
      for (var k = runs.length - 1; k >= 0; k--) {
        safeDeleteRows_(sh, runs[k][0] + 2, runs[k][1] - runs[k][0] + 1);
      }
    }
  }
  if (rows.length) writeRows_(sh, ds, sh.getLastRow() + 1, rows);
  return deleted;
}

/** Записує рядки, ставлячи текстовий формат для текстових колонок (щоб коди "000123" і дати не псувались). */
function writeRows_(sh, ds, startRow, rows) {
  var need = startRow + rows.length - 1;
  if (sh.getMaxRows() < need) sh.insertRowsAfter(sh.getMaxRows(), need - sh.getMaxRows());
  var formats = ds.columns.map(function (c) { return ds.types[c] === 'number' ? '0.###' : '@'; });
  var range = sh.getRange(startRow, 1, rows.length, ds.columns.length);
  range.setNumberFormats(rows.map(function () { return formats; }));
  range.setValues(rows);
}

/** Кількість клітинок у Таблиці (ліміт Google — 10 млн). */
function countCells_(ss) {
  return ss.getSheets().reduce(function (s, sh) { return s + sh.getMaxRows() * sh.getMaxColumns(); }, 0);
}
