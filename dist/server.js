// ЗІБРАНО АВТОМАТИЧНО з src/server (node build.js). Не редагуйте цей файл напряму.
// ===================== Config.js =====================
/**
 * ТРІБО · Дашборд продажів · Проєкт ПРИЙОМУ ДАНИХ (прив'язаний до Google Таблиці)
 * Config.gs — усі налаштування в одному місці.
 *
 * Щоб додати новий набір даних (план, собівартість, оплати…):
 *   1) допишіть його опис у DATASETS нижче;
 *   2) запустіть setup() — аркуш створиться автоматично;
 *   3) програміст 1С надсилає JSON з "dataset": "<назва>".
 * Код прийому міняти не потрібно.
 */

var CONFIG = {
  // Організація за замовчуванням (для резервного імпорту Excel, де її немає у файлі)
  DEFAULT_ORGANIZATION: 'ТОВ "Білоцерківський завод "ТРІБО""',

  // Допуск звірки контрольної суми, грн (через округлення копійок)
  CONTROL_AMOUNT_TOLERANCE: 0.05,

  // Назви службових аркушів
  SHEET_LOG: '_log',
  SHEET_DATASETS: '_datasets',
  SHEET_GROUPS: '_groups',
  SHEET_CONTROL: '_control',

  // Скільки рядків журналу зберігати (старі видаляються)
  LOG_MAX_ROWS: 5000,

  // Папки Google Drive для резервного шляху (створюються setup())
  INBOX_FOLDER_NAME: 'Inbox',
  INBOX_DONE_SUBFOLDER: 'Оброблено',
  INBOX_ERROR_SUBFOLDER: 'Помилки',

  // Межа: якщо Таблиця перевищує цю кількість клітинок — попередження в _log
  CELLS_WARNING: 8000000
};

/**
 * Реєстр наборів даних.
 *  columns    — порядок колонок на аркуші (і обов'язкові поля в JSON-рядку);
 *  types      — 'text' | 'number' | 'month' (перше число місяця, РРРР-ММ-01);
 *  periodField / orgField — за якими полями працює replace_period;
 *  amountField — поле для контрольної суми.
 */
var DATASETS = {
  sales: {
    sheet: 'sales',
    description: 'Продажі (звіт 1С "Продажи"; рахунки 3611, 3621; суми без ПДВ, грн)',
    columns: ['month', 'organization', 'group_path', 'group_ids', 'item_code', 'item', 'unit',
              'counterparty_code', 'counterparty', 'qty', 'amount'],
    types: {
      month: 'month', organization: 'text', group_path: 'text', group_ids: 'text',
      item_code: 'text', item: 'text', unit: 'text',
      counterparty_code: 'text', counterparty: 'text',
      qty: 'number', amount: 'number'
    },
    required: ['month', 'organization', 'item', 'counterparty', 'amount'],
    periodField: 'month',
    orgField: 'organization',
    amountField: 'amount',
    hasProductGroups: true // при записі нові group_path додаються в _groups
  },
  planfact: {
    sheet: 'planfact',
    description: 'План-факт продажів по клієнтах і номенклатурі (звіт 1С, сценарій "Основний щомісячний"), грн без ПДВ',
    columns: ['month', 'organization', 'region', 'counterparty_code', 'counterparty', 'group_path', 'group_ids',
              'item_code', 'item', 'qty_plan', 'qty_fact', 'sum_plan', 'sum_fact', 'sum_fact_vat',
              'cost_plan', 'cost_fact', 'profit_plan', 'profit_fact'],
    types: {
      month: 'month', organization: 'text', region: 'text', counterparty_code: 'text', counterparty: 'text',
      group_path: 'text', group_ids: 'text', item_code: 'text', item: 'text',
      qty_plan: 'number', qty_fact: 'number', sum_plan: 'number', sum_fact: 'number', sum_fact_vat: 'number',
      cost_plan: 'number', cost_fact: 'number', profit_plan: 'number', profit_fact: 'number'
    },
    required: ['month', 'organization', 'counterparty', 'item'],
    periodField: 'month', orgField: 'organization', amountField: 'sum_fact'
  },
  planfact_clients: {
    sheet: 'planfact_clients',
    description: 'Оплати і борг по клієнтах (той самий звіт план-факт, рівень клієнта), грн',
    columns: ['month', 'organization', 'region', 'counterparty_code', 'counterparty', 'payments', 'debt'],
    types: { month: 'month', organization: 'text', region: 'text', counterparty_code: 'text', counterparty: 'text',
             payments: 'number', debt: 'number' },
    required: ['month', 'organization', 'counterparty'],
    periodField: 'month', orgField: 'organization', amountField: 'payments'
  }
  // Приклад майбутнього набору:
  // plan: { sheet: 'plan', description: 'План продажів', columns: [...], types: {...},
  //         required: [...], periodField: 'month', orgField: 'organization', amountField: 'amount' }
};

/**
 * Правило визначення продуктової групи (розділ 9.3 ТЗ).
 * Можна підправити тут; конкретні групи краще виправляти вручну в аркуші _groups.
 */
var GROUP_RULE = {
  // Група з кодом у дужках, не більше 3 рівнів: "(2.1.1) …", "(3.1.2) …", "(2.10) …"
  CODE_REGEX: /^\(\s*\d+(?:\.\d+){0,2}\s*\)/,
  // Запасне правило, якщо в шляху немає груп з кодом (напр. Дафмі)
  FALLBACK_REGEXES: [
    /^колодки\s+(дискового|барабанного)\s+(гальма|тормоза)/i,
    /^(дисковые|барабанные)\s+колодки/i
  ],
  // Відображувані назви за замовчуванням (точний збіг назви групи)
  DISPLAY_OVERRIDES: {
    '(2.1.1) Колодки гальмові МШС': 'МПС',
    '(2.1.1) Колодки тормозные МШС': 'МПС'
  },
  PATH_SEPARATOR: ' > '
};


// ===================== Storage.js =====================
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
      var overlap = v[0] === c.dataset && orgKey_(v[1]) === orgKey_(c.organization) &&
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
      var hit = m >= fromM && m <= toM && orgKey_(oVals[i][0]) === orgKey_(organization);
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
        return !(m2 >= fromM && m2 <= toM && orgKey_(oVals[idx][0]) === orgKey_(organization));
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

/** Порівняння назв організацій без лапок, пробілів і регістру ("ТРІБО"" і "ТРІБО" — та сама). */
function orgKey_(s) { return String(s || '').replace(/["«»'„“”\s]/g, '').toLowerCase(); }

/** Кількість клітинок у Таблиці (ліміт Google — 10 млн). */
function countCells_(ss) {
  return ss.getSheets().reduce(function (s, sh) { return s + sh.getMaxRows() * sh.getMaxColumns(); }, 0);
}


// ===================== Groups.js =====================
/**
 * Groups.gs — продуктові групи (розділ 9.3 ТЗ).
 *
 * Аркуш _groups:  group_path | group_ids | product_group | display_name | status | first_seen
 *   group_path    — повний шлях груп номенклатури з 1С (ключ, унікальний);
 *   product_group — шлях до групи, яка вважається продуктовою (унікальна, бо коди в 1С не унікальні);
 *   display_name  — як показувати на дашборді. Щоб об'єднати групи — дайте їм однакову display_name;
 *   status        — "нова" (призначено правилом), будь-що інше (напр. "перевірено") — ваша позначка.
 * Рядки, що вже є в _groups, скрипт НІКОЛИ не перезаписує — ручні правки зберігаються.
 */

/** Чиста функція: з шляху груп повертає {product_group, name}. */
function resolveProductGroup_(groupPath) {
  var sep = GROUP_RULE.PATH_SEPARATOR;
  var parts = String(groupPath || '').split(sep).map(function (s) { return s.trim(); })
    .filter(function (s) { return s; });
  if (!parts.length) return { product_group: '(без групи)', name: '(без групи)' };

  var idx = -1, i;
  // 1) найглибша група з кодом ≤ 3 рівнів
  for (i = parts.length - 1; i >= 0; i--) {
    if (GROUP_RULE.CODE_REGEX.test(parts[i])) { idx = i; break; }
  }
  // 2) запасне правило ("Колодки дискового/барабанного гальма")
  if (idx < 0) {
    for (i = parts.length - 1; i >= 0 && idx < 0; i--) {
      for (var r = 0; r < GROUP_RULE.FALLBACK_REGEXES.length; r++) {
        if (GROUP_RULE.FALLBACK_REGEXES[r].test(parts[i])) { idx = i; break; }
      }
    }
  }
  // 3) інакше — найглибша група шляху (виправляється вручну в _groups, якщо треба інакше)
  if (idx < 0) idx = parts.length - 1;
  return { product_group: parts.slice(0, idx + 1).join(sep), name: parts[idx] };
}

/** Відображувана назва за замовчуванням; usedNames — {name: product_group} для уникнення збігів. */
function defaultDisplayName_(resolved, usedNames) {
  var name = GROUP_RULE.DISPLAY_OVERRIDES[resolved.name] || resolved.name;
  if (usedNames && usedNames[name] && usedNames[name] !== resolved.product_group) {
    var parts = resolved.product_group.split(GROUP_RULE.PATH_SEPARATOR);
    var parent = parts.length > 1 ? parts[parts.length - 2] : '';
    name = name + (parent ? ' (' + parent + ')' : ' (2)');
  }
  return name;
}

/** Додає в _groups нові шляхи груп. pathsToIds — {group_path: group_ids}. */
function syncGroups_(ss, pathsToIds) {
  var sh = getOrCreateSheet_(ss, CONFIG.SHEET_GROUPS,
    ['group_path', 'group_ids', 'product_group', 'display_name', 'status', 'first_seen']);
  var last = sh.getLastRow();
  var existing = {}, usedNames = {}, pgToName = {};
  if (last > 1) {
    sh.getRange(2, 1, last - 1, 4).getValues().forEach(function (r) {
      existing[String(r[0])] = true;
      usedNames[String(r[3])] = String(r[2]);
      pgToName[String(r[2])] = String(r[3]);
    });
  }
  var now = nowIso_();
  var add = [];
  Object.keys(pathsToIds).forEach(function (p) {
    if (existing[p]) return;
    var res = resolveProductGroup_(p);
    var dn = pgToName[res.product_group] || defaultDisplayName_(res, usedNames);
    usedNames[dn] = res.product_group;
    pgToName[res.product_group] = dn;
    existing[p] = true;
    add.push([p, pathsToIds[p] || '', res.product_group, dn, 'нова', now]);
  });
  if (add.length) {
    var start = sh.getLastRow() + 1, need = start + add.length - 1;
    if (need > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), need - sh.getMaxRows());
    var range = sh.getRange(start, 1, add.length, 6);
    range.setNumberFormat('@');
    range.setValues(add);
  }
  return add.length;
}

/**
 * Перебудовує довідник _groups з нуля за шляхами груп, що є в аркуші sales.
 * Запускати, якщо в _groups лишилися тестові групи. УВАГА: ручні правки в _groups буде втрачено.
 */
function rebuildGroups() {
  var ss = getSpreadsheet_();
  var ds = DATASETS.sales, sh = ss.getSheetByName(ds.sheet);
  var gsh = ss.getSheetByName(CONFIG.SHEET_GROUPS);
  if (gsh && gsh.getLastRow() > 1) safeDeleteRows_(gsh, 2, gsh.getLastRow() - 1);
  var paths = {};
  if (sh && sh.getLastRow() > 1) {
    var pc = ds.columns.indexOf('group_path') + 1, ic = ds.columns.indexOf('group_ids') + 1;
    var p = sh.getRange(2, pc, sh.getLastRow() - 1, 1).getValues();
    var ids = sh.getRange(2, ic, sh.getLastRow() - 1, 1).getValues();
    p.forEach(function (r, i) { var k = String(r[0]); if (k && !(k in paths)) paths[k] = String(ids[i][0]); });
  }
  var n = syncGroups_(ss, paths);
  updateDatasetRegistry_(ss, 'sales', DATASETS.sales, { updated: nowIso_(), generated_at: '' }); // скинути кеш дашборду
  log_(ss, { source: 'manual', dataset: 'sales', status: 'OK', message: 'Довідник _groups перебудовано: ' + n + ' шляхів груп' });
  console.log('Довідник груп перебудовано: ' + n + ' шляхів.');
}


// ===================== Ingest.js =====================
/**
 * Ingest.gs — прийом даних з 1С (HTTPS POST, JSON).
 *
 * Відповідь — JSON: {"ok": true, "rows": 1234, "amount": 9790198.22, ...}
 *                або {"ok": false, "error": "..."}
 * Увага: Google відповідає на POST редиректом 302; тіло відповіді (цей JSON)
 * лежить за адресою з заголовка Location і читається GET-запитом.
 */

function doPost(e) {
  var t0 = Date.now();
  var payload;
  try {
    var body = e && e.postData ? e.postData.contents : '';
    payload = JSON.parse(body);
  } catch (err) {
    var ss0 = getSpreadsheet_();
    log_(ss0, { source: 'http', status: 'ERROR', message: 'Некоректний JSON: ' + err, duration_ms: Date.now() - t0 });
    return json_({ ok: false, error: 'Некоректний JSON: ' + err });
  }

  var token = PropertiesService.getScriptProperties().getProperty('TOKEN');
  if (!token || payload.token !== token) {
    log_(getSpreadsheet_(), { source: 'http', dataset: payload.dataset, status: 'DENIED',
      message: 'Невірний токен', duration_ms: Date.now() - t0 });
    return json_({ ok: false, error: 'Невірний токен' });
  }
  return json_(ingestPayload_(payload, 'http'));
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/**
 * Спільне ядро для HTTP, імпорту Excel і тестів. Токен тут уже не перевіряється.
 * source — 'http' | 'excel' | 'test'
 */
function ingestPayload_(payload, source) {
  var t0 = Date.now();
  var ss = getSpreadsheet_();
  var period = (payload && payload.period_from) + '…' + (payload && payload.period_to);
  var logBase = { source: source, dataset: payload && payload.dataset, organization: payload && payload.organization,
    period: period, mode: payload && payload.mode };

  var v = validatePayload_(payload);
  if (v.errors.length) {
    var msg = v.errors.slice(0, 10).join('; ') + (v.errors.length > 10 ? ' … (+' + (v.errors.length - 10) + ')' : '');
    log_(ss, merge_(logBase, { rows: v.rowCount, amount: v.amount, status: 'ERROR', message: msg, duration_ms: Date.now() - t0 }));
    return { ok: false, error: msg, rows: v.rowCount, amount: v.amount };
  }

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(60000)) {
    log_(ss, merge_(logBase, { status: 'ERROR', message: 'Зайнято іншим завантаженням (lock)', duration_ms: Date.now() - t0 }));
    return { ok: false, error: 'Зайнято іншим завантаженням, повторіть пізніше' };
  }
  try {
    var ds = DATASETS[payload.dataset];
    var deleted = replacePeriod_(ss, ds, payload.organization, payload.period_from, payload.period_to, v.rows);

    var newGroups = 0;
    if (ds.hasProductGroups) newGroups = syncGroups_(ss, v.groupPaths);

    saveControl_(ss, { dataset: payload.dataset, organization: payload.organization,
      period_from: payload.period_from, period_to: payload.period_to,
      rows: payload.control ? payload.control.rows : null,
      amount: payload.control ? payload.control.amount : null, source: source });

    var updated = nowIso_();
    updateDatasetRegistry_(ss, payload.dataset, ds, { updated: updated, generated_at: payload.generated_at || '' });

    var cells = countCells_(ss);
    var note = 'видалено ' + deleted + ', записано ' + v.rows.length +
      (newGroups ? ', нових груп: ' + newGroups + ' (перевірте _groups)' : '');
    if (cells > CONFIG.CELLS_WARNING) note += '; УВАГА: ' + cells + ' клітинок — час винести архівні роки';
    log_(ss, merge_(logBase, { rows: v.rows.length, amount: round2_(v.amount), status: 'OK', message: note,
      duration_ms: Date.now() - t0 }));
    return { ok: true, dataset: payload.dataset, rows: v.rows.length, amount: round2_(v.amount),
      deleted: deleted, new_groups: newGroups, updated: updated };
  } catch (err) {
    log_(ss, merge_(logBase, { status: 'ERROR', message: 'Помилка запису: ' + err, duration_ms: Date.now() - t0 }));
    return { ok: false, error: 'Помилка запису: ' + err };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Чиста функція перевірки. Повертає {errors, rows (масиви у порядку колонок), amount, rowCount, groupPaths}.
 */
function validatePayload_(p) {
  var errors = [], out = [], sum = 0, groupPaths = {};
  var res = function () { return { errors: errors, rows: out, amount: round2_(sum), rowCount: out.length, groupPaths: groupPaths }; };
  if (!p || typeof p !== 'object') { errors.push('Порожній запит'); return res(); }

  var ds = DATASETS[p.dataset];
  if (!ds) { errors.push('Невідомий dataset "' + p.dataset + '". Відомі: ' + Object.keys(DATASETS).join(', ')); return res(); }
  if (!p.organization) errors.push('Не вказано organization');
  if (!isIsoDate_(p.period_from)) errors.push('period_from має бути РРРР-ММ-ДД');
  if (!isIsoDate_(p.period_to)) errors.push('period_to має бути РРРР-ММ-ДД');
  if (p.period_from > p.period_to) errors.push('period_from пізніше за period_to');
  if (p.mode !== 'replace_period') errors.push('Підтримується лише mode = "replace_period"');
  if (!Array.isArray(p.rows)) errors.push('rows має бути масивом');
  if (!p.control || typeof p.control.rows !== 'number' || typeof p.control.amount !== 'number') {
    errors.push('control.rows і control.amount обов\'язкові (числа)');
  }
  if (errors.length) return res();

  var fromM = p.period_from.slice(0, 7), toM = p.period_to.slice(0, 7);
  for (var i = 0; i < p.rows.length; i++) {
    var r = p.rows[i] || {};
    if (r.organization === undefined || r.organization === null || r.organization === '') r.organization = p.organization;
    var arr = [];
    for (var c = 0; c < ds.columns.length; c++) {
      var col = ds.columns[c], type = ds.types[col], val = r[col];
      if (val === undefined || val === null) val = '';
      if (type === 'number') {
        if (val === '') val = 0;
        if (typeof val === 'string' && /^-?\d+(\.\d+)?$/.test(val.trim())) val = parseFloat(val);
        if (typeof val !== 'number' || !isFinite(val)) { errors.push('Рядок ' + (i + 1) + ': ' + col + ' не число'); val = 0; }
      } else if (type === 'month') {
        if (!isIsoDate_(val)) { errors.push('Рядок ' + (i + 1) + ': ' + col + ' має бути РРРР-ММ-ДД'); }
        else {
          val = String(val).slice(0, 7) + '-01';
          if (val.slice(0, 7) < fromM || val.slice(0, 7) > toM) errors.push('Рядок ' + (i + 1) + ': ' + col + ' поза періодом');
        }
      } else {
        val = String(val).trim();
      }
      arr.push(val);
    }
    for (var q = 0; q < ds.required.length; q++) {
      var f = ds.required[q];
      if (arr[ds.columns.indexOf(f)] === '') errors.push('Рядок ' + (i + 1) + ': порожнє ' + f);
    }
    if (arr[ds.columns.indexOf(ds.orgField)] !== p.organization) errors.push('Рядок ' + (i + 1) + ': інша організація');
    sum += arr[ds.columns.indexOf(ds.amountField)];
    if (ds.hasProductGroups) {
      var gp = arr[ds.columns.indexOf('group_path')];
      if (gp && !(gp in groupPaths)) groupPaths[gp] = arr[ds.columns.indexOf('group_ids')];
    }
    out.push(arr);
  }

  if (p.control.rows !== out.length) errors.push('Контроль: рядків ' + out.length + ', очікувалось ' + p.control.rows);
  if (Math.abs(round2_(sum) - p.control.amount) > CONFIG.CONTROL_AMOUNT_TOLERANCE) {
    errors.push('Контроль: сума ' + round2_(sum) + ', очікувалось ' + p.control.amount);
  }
  return res();
}

function isIsoDate_(s) { return typeof s === 'string' && /^\d{4}-(0[1-9]|1[0-2])-(0[1-9]|[12]\d|3[01])$/.test(s); }
function round2_(x) { return Math.round(x * 100) / 100; }
function merge_(a, b) { var o = {}, k; for (k in a) o[k] = a[k]; for (k in b) o[k] = b[k]; return o; }


// ===================== ExcelImport.js =====================
/**
 * ExcelImport.gs — резервний шлях: Excel-звіт 1С "Продажи" → набір sales.
 *
 * Покладіть .xlsx у папку Google Диска "Inbox". Кожні 15 хв (або з меню "ТРІБО")
 * скрипт розбирає файл, записує кожен місяць з файлу через replace_period
 * і переносить файл у "Inbox/Оброблено" (або "Inbox/Помилки" — причина в _log).
 *
 * Розбір написано без сторонніх бібліотек (xlsx = zip з XML), тому він не залежить від
 * регістру імен файлів усередині (xl/SharedStrings.xml з великої літери у 1С).
 */

var EXCEL = {
  QTY_HEADER: /^(количество|кількість)$/i,
  AMOUNT_HEADER: /^сумма|^сума/i,
  TOTAL_LABEL: /^(итог|итого|разом|всього|усього)\.?$/i,
  // вузли ієрархії, які треба пропустити (напр. рядок з назвою організації, якщо звіт згруповано по ній)
  SKIP_NODES: [/білоцерківський завод|белоцерковский завод/i],
  // що вважати одиницею виміру після останньої коми у назві номенклатури
  UNIT_REGEX: /^[A-Za-zА-Яа-яІіЇїЄєҐґ][A-Za-zА-Яа-яІіЇїЄєҐґ.\/ ²³0-9-]{0,14}$/
};

var MONTHS_ = {
  'январ': 1, 'феврал': 2, 'март': 3, 'апрел': 4, 'ма': 5, 'июн': 6, 'июл': 7, 'август': 8, 'сентябр': 9, 'октябр': 10, 'ноябр': 11, 'декабр': 12,
  'січ': 1, 'лют': 2, 'берез': 3, 'квіт': 4, 'травн': 5, 'черв': 6, 'лип': 7, 'серп': 8, 'верес': 9, 'жовт': 10, 'листопад': 11, 'груд': 12
};

/* ===================== Тригер / Drive ===================== */

function processInbox() {
  var props = PropertiesService.getScriptProperties();
  var folderId = props.getProperty('INBOX_FOLDER_ID');
  if (!folderId) { console.log('Запустіть setup()'); return; }
  var inbox = DriveApp.getFolderById(folderId);
  var done = getOrCreateFolder_(inbox, CONFIG.INBOX_DONE_SUBFOLDER);
  var bad = getOrCreateFolder_(inbox, CONFIG.INBOX_ERROR_SUBFOLDER);
  var ss = getSpreadsheet_();
  var started = Date.now();

  var files = inbox.getFiles();
  while (files.hasNext()) {
    if (Date.now() - started > 4 * 60 * 1000) break; // запас до ліміту 6 хв; решта — наступного разу
    var f = files.next();
    var name = f.getName();
    if (!/\.xlsx$/i.test(name)) {
      log_(ss, { source: 'excel', status: 'ERROR', message: name + ': потрібен формат .xlsx (у 1С: Зберегти як → Лист Excel 2007)' });
      f.moveTo(bad);
      continue;
    }
    try {
      var result = importSalesXlsxBlob_(f.getBlob(), name);
      f.moveTo(result.ok ? done : bad);
    } catch (err) {
      log_(ss, { source: 'excel', dataset: 'sales', status: 'ERROR', message: name + ': ' + err });
      f.moveTo(bad);
    }
  }
}

function importSalesXlsxBlob_(blob, fileName) {
  var ss = getSpreadsheet_();
  var copy = blob.copyBlob().setContentType('application/zip');
  var parts = Utilities.unzip(copy);
  var files = {};
  parts.forEach(function (b) { files[b.getName()] = b; });
  var table = readXlsxTable_(function (path) {
    var key = findKeyCaseInsensitive_(files, path);
    return key ? files[key].getDataAsString('UTF-8') : null;
  }, Object.keys(files));

  // тип звіту: «План-факт» (є стовпці план/факт) або «Продажи»
  var isPF = isPlanFactTable_(table);
  var dsName = isPF ? 'planfact' : 'sales';
  var parsed = isPF
    ? parsePlanFactTable_(table, { organization: CONFIG.DEFAULT_ORGANIZATION })
    : parse1cSalesTable_(table, { organization: CONFIG.DEFAULT_ORGANIZATION, fallbackYear: new Date().getFullYear() });
  if (parsed.errors.length) {
    log_(ss, { source: 'excel', dataset: dsName, status: 'ERROR', message: fileName + ': ' + parsed.errors.join('; ') });
    return { ok: false };
  }
  var allOk = true;
  parsed.payloads.forEach(function (p) {
    p.generated_at = nowIso_();
    var r = ingestPayload_(p, 'excel');
    if (!r.ok) allOk = false;
  });
  if (parsed.warnings.length) {
    log_(ss, { source: 'excel', dataset: dsName, status: 'WARN', message: fileName + ': ' + parsed.warnings.join('; ') });
  }
  return { ok: allOk };
}

/* ===================== Чисті функції (без Google API) ===================== */

function findKeyCaseInsensitive_(obj, path) {
  var low = path.toLowerCase();
  for (var k in obj) if (k.toLowerCase() === low) return k;
  return null;
}

function xmlDecode_(s) {
  return s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'")
    .replace(/&#x([0-9a-f]+);/gi, function (m, h) { return String.fromCharCode(parseInt(h, 16)); })
    .replace(/&#(\d+);/g, function (m, d) { return String.fromCharCode(parseInt(d, 10)); })
    .replace(/&amp;/g, '&');
}

function textOfT_(xml) {
  xml = xml.replace(/<rPh\b[\s\S]*?<\/rPh>/g, '');
  var out = '', m, re = /<t\b[^>]*>([\s\S]*?)<\/t>/g;
  while ((m = re.exec(xml))) out += m[1];
  return xmlDecode_(out);
}

function colIndex_(ref) {
  var letters = String(ref).replace(/[^A-Z]/gi, '').toUpperCase(), n = 0;
  for (var i = 0; i < letters.length; i++) n = n * 26 + (letters.charCodeAt(i) - 64);
  return n - 1;
}

/**
 * readFile(path) → рядок XML або null. Повертає [{r, level, cells: []}] для першого аркуша.
 */
function readXlsxTable_(readFile, allNames) {
  // спільні рядки — у 1С файл може називатися SharedStrings.xml
  var shared = [];
  var ssXml = readFile('xl/sharedStrings.xml');
  if (ssXml) {
    var m, re = /<si\b[^>]*>([\s\S]*?)<\/si>|<si\s*\/>/g;
    while ((m = re.exec(ssXml))) shared.push(m[1] ? textOfT_(m[1]) : '');
  }
  // перший аркуш книги
  var sheetPath = null;
  var wb = readFile('xl/workbook.xml'), rels = readFile('xl/_rels/workbook.xml.rels');
  if (wb && rels) {
    var sm = /<sheet\b[^>]*\br:id="([^"]+)"/.exec(wb) || /<sheet\b[^>]*\bid="([^"]+)"/.exec(wb);
    if (sm) {
      var relRe = new RegExp('<Relationship\\b[^>]*\\bId="' + sm[1] + '"[^>]*>', 'i');
      var rm = relRe.exec(rels);
      if (rm) {
        var t = /Target="([^"]+)"/.exec(rm[0]);
        if (t) sheetPath = t[1].replace(/^\/?xl\//i, '').replace(/^\//, '');
        if (sheetPath) sheetPath = 'xl/' + sheetPath;
      }
    }
  }
  var sheetXml = sheetPath ? readFile(sheetPath) : null;
  if (!sheetXml) {
    var cand = (allNames || []).filter(function (n) { return /^xl\/worksheets\/sheet\d*\.xml$/i.test(n); }).sort();
    if (cand.length) sheetXml = readFile(cand[0]);
  }
  if (!sheetXml) throw new Error('Не знайдено аркуш у файлі xlsx');

  var rows = [], rm2, rowRe = /<row\b([^>]*?)(?:\/>|>([\s\S]*?)<\/row>)/g, autoR = 0;
  while ((rm2 = rowRe.exec(sheetXml))) {
    var attrs = rm2[1], body = rm2[2] || '';
    var rAttr = /\br="(\d+)"/.exec(attrs);
    var rNum = rAttr ? parseInt(rAttr[1], 10) : autoR + 1;
    autoR = rNum;
    var lv = /\boutlineLevel="(\d+)"/.exec(attrs);
    var cells = [], cm, cellRe = /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g, autoC = -1;
    while ((cm = cellRe.exec(body))) {
      var ca = cm[1], cb = cm[2] || '';
      var ref = /\br="([A-Z]+)\d*"/i.exec(ca);
      var ci = ref ? colIndex_(ref[1]) : autoC + 1;
      autoC = ci;
      var type = (/\bt="([^"]+)"/.exec(ca) || [])[1] || 'n';
      var val = null;
      if (type === 'inlineStr') {
        var is = /<is>([\s\S]*?)<\/is>/.exec(cb);
        val = is ? textOfT_(is[1]) : '';
      } else {
        var v = /<v>([\s\S]*?)<\/v>/.exec(cb);
        if (v) {
          if (type === 's') val = shared[parseInt(v[1], 10)];
          else if (type === 'str' || type === 'e') val = xmlDecode_(v[1]);
          else if (type === 'b') val = v[1] === '1';
          else val = parseFloat(v[1]);
        }
      }
      if (val !== null && val !== undefined && val !== '') cells[ci] = val;
    }
    rows.push({ r: rNum, level: lv ? parseInt(lv[1], 10) : 0, cells: cells });
  }
  return rows;
}

function toNumber_(v) {
  if (typeof v === 'number') return v;
  if (v === null || v === undefined) return 0;
  var s = String(v).replace(/[\s ]/g, '').replace(',', '.');
  var n = parseFloat(s);
  return isFinite(n) ? n : 0;
}

/** "Январь 2026" / "Січень 2026 р." / "01.2026" / "2026-01" → {m, y|null} */
function parseMonthLabel_(label) {
  var s = String(label || '').trim().toLowerCase();
  var m = /^(\d{1,2})[.\/](\d{4})$/.exec(s);
  if (m) return { m: parseInt(m[1], 10), y: parseInt(m[2], 10) };
  m = /^(\d{4})-(\d{1,2})$/.exec(s);
  if (m) return { m: parseInt(m[2], 10), y: parseInt(m[1], 10) };
  var word = (/^[a-zа-яіїєґ']+/.exec(s) || [''])[0];
  if (!word) return null;
  var best = null, bestLen = 0;
  for (var k in MONTHS_) {
    if (word.indexOf(k) === 0 && k.length > bestLen) { best = MONTHS_[k]; bestLen = k.length; }
  }
  if (!best) return null;
  if (best === 5 && !/^ма[йя]/.test(word)) return null; // "ма" — лише "май/мая"
  var y = /(\d{4})/.exec(s);
  return { m: best, y: y ? parseInt(y[1], 10) : null };
}

/**
 * Розбирає таблицю звіту 1С "Продажи".
 * Повертає {payloads: [JSON як від 1С, по одному на місяць], errors: [], warnings: []}.
 */
function parse1cSalesTable_(rows, opts) {
  var errors = [], warnings = [];
  var org = opts.organization;

  // 1) рядок з "Количество" / "Сумма…"
  var hIdx = -1;
  for (var i = 0; i < rows.length && i < 60; i++) {
    var cnt = 0;
    rows[i].cells.forEach(function (v) { if (typeof v === 'string' && EXCEL.QTY_HEADER.test(v.trim())) cnt++; });
    if (cnt >= 1) { hIdx = i; break; }
  }
  if (hIdx < 0) return { payloads: [], errors: ['не знайдено рядок заголовків з "Количество"'], warnings: warnings };
  var hdr = rows[hIdx].cells;

  // порядок групувань із шапки звіту: "Группировки строк: Номенклатура (Ієрархія); Контрагент (Елементи);"
  var order = null;
  for (var g0 = 0; g0 < hIdx && !order; g0++) {
    rows[g0].cells.forEach(function (v) {
      var gm = /группировки строк:\s*([^\n]*)/i.exec(String(v));
      if (gm && !order) {
        var ni = gm[1].search(/номенклатур/i), ki = gm[1].search(/контрагент/i);
        if (ni >= 0 && ki >= 0) order = ni < ki ? 'item_first' : 'counterparty_first';
      }
    });
  }

  // рік з шапки ("Период: 01.01.2026 - 11.11.2026"), якщо в назвах місяців року немає
  var headerYear = null;
  for (var h = 0; h < hIdx && headerYear === null; h++) {
    rows[h].cells.forEach(function (v) {
      var mm = /(\d{2})\.(\d{2})\.(\d{4})/.exec(String(v));
      if (mm && headerYear === null) headerYear = parseInt(mm[3], 10);
    });
  }

  // 2) пари колонок місяців
  var monthCols = [], totalCols = null;
  for (var c = 0; c < hdr.length; c++) {
    if (typeof hdr[c] !== 'string' || !EXCEL.QTY_HEADER.test(hdr[c].trim())) continue;
    var amtCol = c + 1;
    if (!(typeof hdr[amtCol] === 'string' && EXCEL.AMOUNT_HEADER.test(hdr[amtCol].trim()))) {
      warnings.push('колонка ' + (c + 2) + ': очікувалась "Сумма…"');
    }
    var label = null;
    for (var up = hIdx - 1; up >= Math.max(0, hIdx - 4) && label === null; up--) {
      var cl = rows[up].cells;
      if (cl[c] !== undefined) label = cl[c];
      else if (cl[amtCol] !== undefined) label = cl[amtCol];
    }
    if (label !== null && EXCEL.TOTAL_LABEL.test(String(label).trim())) { totalCols = { qty: c, amount: amtCol }; continue; }
    var pm = parseMonthLabel_(label);
    if (!pm) { warnings.push('не розпізнано місяць "' + label + '" (колонка ' + (c + 1) + ')'); continue; }
    var year = pm.y || headerYear || opts.fallbackYear;
    monthCols.push({ month: year + '-' + ('0' + pm.m).slice(-2) + '-01', qty: c, amount: amtCol });
  }
  if (!monthCols.length) errors.push('не знайдено колонок місяців');

  // колонка з назвами (ієрархія) — найлівіша колонка з текстом у даних
  var firstNum = monthCols.length ? monthCols[0].qty : 1;
  var labelCol = 0;
  for (var lc = 0; lc < firstNum; lc++) {
    var hasText = false;
    for (var q = hIdx + 1; q < Math.min(rows.length, hIdx + 40); q++) if (typeof rows[q].cells[lc] === 'string') { hasText = true; break; }
    if (hasText) { labelCol = lc; break; }
  }

  // 3) дані: розгортання ієрархії, лише найнижчий рівень
  var data = [], totalRow = null;
  for (var d = hIdx + 1; d < rows.length; d++) {
    var lab = rows[d].cells[labelCol];
    if (lab === undefined || lab === null || String(lab).trim() === '') continue;
    lab = String(lab).trim();
    if (EXCEL.TOTAL_LABEL.test(lab)) { totalRow = rows[d]; continue; }
    if (/^(номенклатура|контрагент|группа|група)/i.test(lab) && d === hIdx + 1) continue; // підзаголовок
    data.push({ level: rows[d].level, label: lab, cells: rows[d].cells });
  }
  var stack = [], leaves = [];
  for (var j = 0; j < data.length; j++) {
    var row = data[j];
    stack.length = row.level;
    stack[row.level] = row.label;
    var next = data[j + 1];
    if (!next || next.level <= row.level) {
      leaves.push({ path: stack.filter(function (x) { return x !== undefined; }), cells: row.cells });
    }
  }
  if (!leaves.length) errors.push('не знайдено рядків з даними');

  // 4) класифікація вузлів шляху: групи (закінчуються комою), номенклатура (", од."), контрагент
  var outByMonth = {};
  monthCols.forEach(function (mc) { outByMonth[mc.month] = []; });
  leaves.forEach(function (lf) {
    var nodes = lf.path.filter(function (n) {
      return !EXCEL.SKIP_NODES.some(function (re) { return re.test(n); });
    });
    var info = order ? classifyByOrder_(nodes, order) : classifyPath_(nodes);
    if (info.error) { warnings.push(info.error + ': ' + nodes.join(' > ')); }
    monthCols.forEach(function (mc) {
      var qty = toNumber_(lf.cells[mc.qty]), amt = toNumber_(lf.cells[mc.amount]);
      if (!qty && !amt) return;
      outByMonth[mc.month].push({
        month: mc.month, organization: org,
        group_path: info.groups.join(GROUP_RULE.PATH_SEPARATOR),
        group_ids: '', // в Excel ідентифікаторів 1С немає
        item_code: '', item: info.item, unit: info.unit,
        counterparty_code: '', counterparty: info.counterparty,
        qty: Math.round(qty * 1000) / 1000, amount: Math.round(amt * 100) / 100
      });
    });
  });

  // 5) пакети по місяцях + контроль з рядка "Итог"
  var payloads = monthCols.map(function (mc) {
    var rowsM = outByMonth[mc.month];
    var sum = rowsM.reduce(function (s, r) { return s + r.amount; }, 0);
    var control = sum;
    if (totalRow) control = toNumber_(totalRow.cells[mc.amount]);
    else warnings.push('немає рядка "Итог" — контрольна сума не перевіряється');
    var lastDay = new Date(Date.UTC(parseInt(mc.month.slice(0, 4), 10), parseInt(mc.month.slice(5, 7), 10), 0)).getUTCDate();
    return {
      dataset: 'sales', organization: org,
      period_from: mc.month, period_to: mc.month.slice(0, 8) + ('0' + lastDay).slice(-2),
      mode: 'replace_period',
      control: { rows: rowsM.length, amount: Math.round(control * 100) / 100 },
      rows: rowsM
    };
  });
  if (totalRow && totalCols) {
    var gt = toNumber_(totalRow.cells[totalCols.amount]);
    var sumMonths = payloads.reduce(function (s, p) { return s + p.control.amount; }, 0);
    if (Math.abs(gt - sumMonths) > 0.05) warnings.push('"Итог" загальний ' + gt + ' ≠ сума місяців ' + Math.round(sumMonths * 100) / 100);
  }
  return { payloads: payloads, errors: errors, warnings: dedupe_(warnings).slice(0, 20) };
}

/** Класифікація за відомим порядком групувань (надійніше: назви контрагентів теж можуть закінчуватися комою). */
function classifyByOrder_(nodes, order) {
  var clean = function (n) { return n.replace(/,\s*$/, '').trim(); };
  if (nodes.length < 2) return classifyPath_(nodes);
  var itemNode, cp, groups;
  if (order === 'item_first') { // групи > номенклатура > контрагент
    cp = nodes[nodes.length - 1]; itemNode = nodes[nodes.length - 2]; groups = nodes.slice(0, -2);
  } else {                      // контрагент > групи > номенклатура
    cp = nodes[0]; itemNode = nodes[nodes.length - 1]; groups = nodes.slice(1, -1);
  }
  var su = splitUnit_(itemNode);
  return { groups: groups.map(clean), item: su.name, unit: su.unit, counterparty: clean(cp) || '(без контрагента)', error: null };
}

function classifyPath_(nodes) {
  var isGroup = function (n) { return /,\s*$/.test(n); };
  var gFirst = -1, gLast = -1;
  nodes.forEach(function (n, i) { if (isGroup(n)) { if (gFirst < 0) gFirst = i; gLast = i; } });
  var groups = gFirst < 0 ? [] : nodes.slice(gFirst, gLast + 1).filter(isGroup).map(function (n) { return n.replace(/,\s*$/, '').trim(); });
  var itemIdx = -1, cp = [];
  if (gLast >= 0 && gLast + 1 < nodes.length) {
    itemIdx = gLast + 1;
  } else {
    for (var i = nodes.length - 1; i >= 0; i--) if (!isGroup(nodes[i]) && splitUnit_(nodes[i]).unit) { itemIdx = i; break; }
  }
  if (itemIdx < 0) return { groups: groups, item: '(невідомо)', unit: '', counterparty: nodes.filter(function (n) { return !isGroup(n); }).join(' / ') || '(невідомо)', error: 'не знайдено номенклатуру' };
  nodes.forEach(function (n, i) { if (i !== itemIdx && !isGroup(n)) cp.push(n); });
  var su = splitUnit_(nodes[itemIdx]);
  return { groups: groups, item: su.name, unit: su.unit, counterparty: cp.join(' / ') || '(без контрагента)', error: cp.length ? null : 'не знайдено контрагента' };
}

/** "2ТР-155, тис.шт" → {name: "2ТР-155", unit: "тис.шт"} */
function splitUnit_(s) {
  var k = s.lastIndexOf(',');
  if (k < 0) return { name: s.trim(), unit: '' };
  var u = s.slice(k + 1).trim();
  if (!EXCEL.UNIT_REGEX.test(u)) return { name: s.trim(), unit: '' };
  return { name: s.slice(0, k).trim(), unit: u };
}

function dedupe_(arr) { var seen = {}; return arr.filter(function (x) { return seen[x] ? false : (seen[x] = true); }); }


// ===================== Setup.js =====================
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


// ===================== TestData.js =====================
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


// ===================== Dashboard.js =====================
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
    yearPeriod: getPeriodCached_(year, 0, y),
    planfactMonths: apiPlanFactMonths()
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
    // версія = час оновлення всіх наборів (кеш скидається, щойно оновився будь-який)
    upd = ds.getRange(2, 1, ds.getLastRow() - 1, 5).getDisplayValues().map(function (r) { return r[0] + '=' + r[4]; }).join(';');
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
  var um = /(?:^|;)sales=([^;#]*)/.exec(version || ''), updated = um ? um[1] : '';

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



// ===================== PlanFact.js =====================
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


// ===================== Auth.js =====================
/**
 * Auth.js — вхід у дашборд за кодом із пошти (без входу в Google).
 *
 * Хто може входити — аркуш "_users" у Таблиці: email | name | active (так/ні) | added | password.
 * Щоб додати людину — допишіть рядок: пошта, ім'я, "так", (дата — необов'язково), пароль.
 * Щоб забрати доступ — поставте "ні". Вхід: пошта + пароль; якщо пароля немає — код на пошту.
 * Сесія живе AUTH.SESSION_DAYS днів; зберігається в Script Properties як хеш токена.
 */

var AUTH = {
  USERS_SHEET: '_users',
  CODE_TTL_SEC: 600,        // код дійсний 10 хв
  MAX_CODES_PER_HOUR: 5,    // захист від розсилки
  MAX_ATTEMPTS: 5,          // спроб ввести код
  SESSION_DAYS: 90,
  MAX_PASSWORD_TRIES: 10   // за 15 хв
};

var AUTH_REQUIRED = 'AUTH_REQUIRED';

function normEmail_(e) { return String(e || '').trim().toLowerCase(); }

function sha_(s) {
  return Utilities.base64EncodeWebSafe(
    Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(s), Utilities.Charset.UTF_8)).replace(/=+$/, '');
}

/**
 * Гарантує, що аркуш _users має 5 стовпців (email | name | active | added | password)
 * і що власник у ньому є. Викликається при відкритті сторінки.
 */
function ensureUsersSheet_() {
  var ss = getSpreadsheet_();
  var sh = getOrCreateSheet_(ss, AUTH.USERS_SHEET, ['email', 'name', 'active', 'added', 'password']);
  if (sh.getMaxColumns() < 5) sh.insertColumnsAfter(sh.getMaxColumns(), 5 - sh.getMaxColumns());
  var head = sh.getRange(1, 1, 1, 5).getValues()[0];
  var want = ['email', 'name', 'active', 'added', 'password'];
  if (want.some(function (h, i) { return String(head[i]) !== h; })) {
    sh.getRange(1, 1, 1, 5).setValues([want]).setFontWeight('bold');
  }
  if (sh.getLastRow() < 2) {
    var owner = '';
    try { owner = Session.getEffectiveUser().getEmail(); } catch (e) {}
    if (owner) {
      if (sh.getMaxRows() < 2) sh.insertRowsAfter(sh.getMaxRows(), 1);
      sh.getRange(2, 1, 1, 4).setValues([[owner, 'Власник', 'так', nowIso_()]]);
    }
  }
  return sh;
}

/** {email: {name, password}} для активних користувачів. */
function usersTable_() {
  var sh = getSpreadsheet_().getSheetByName(AUTH.USERS_SHEET), out = {};
  if (!sh || sh.getLastRow() < 2) return out;
  if (sh.getMaxColumns() < 5) return usersFrom_(sh.getRange(2, 1, sh.getLastRow() - 1, sh.getMaxColumns()).getDisplayValues());
  return usersFrom_(sh.getRange(2, 1, sh.getLastRow() - 1, 5).getDisplayValues());
}

function usersFrom_(rows) {
  var out = {};
  rows.forEach(function (r) {
    var e = normEmail_(r[0]), active = String(r[2] || '').trim().toLowerCase();
    if (e && active !== 'ні' && active !== 'no' && active !== 'нет') out[e] = { name: String(r[1] || ''), password: String(r[4] || '') };
  });
  return out;
}

/** {email: name} для активних користувачів. */
function allowedUsers_() {
  var t = usersTable_(), out = {};
  Object.keys(t).forEach(function (e) { out[e] = t[e].name; });
  return out;
}

function randomCode_() {
  var hex = Utilities.getUuid().replace(/-/g, '');
  return ('000000' + (parseInt(hex.slice(0, 12), 16) % 1000000)).slice(-6);
}

/** Крок 1: надіслати код. Відповідь однакова для дозволених і недозволених адрес. */
function authRequestCode_(email) {
  email = normEmail_(email);
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) throw new Error('Перевірте адресу пошти');
  var cache = CacheService.getScriptCache();
  var rlKey = 'rl:' + sha_(email), n = parseInt(cache.get(rlKey) || '0', 10);
  if (n >= AUTH.MAX_CODES_PER_HOUR) throw new Error('Забагато запитів коду. Спробуйте за годину.');
  cache.put(rlKey, String(n + 1), 3600);

  var users = allowedUsers_();
  if (users[email] !== undefined) {
    var code = randomCode_();
    cache.put('code:' + sha_(email), JSON.stringify({ h: sha_(email + ':' + code), a: 0 }), AUTH.CODE_TTL_SEC);
    MailApp.sendEmail({
      to: email,
      subject: 'Код входу в дашборд ТРІБО: ' + code,
      body: 'Ваш код входу в дашборд продажів ТРІБО: ' + code + '\n\nКод дійсний 10 хвилин. ' +
        'Якщо ви не намагалися увійти — просто проігноруйте цей лист.'
    });
    log_(getSpreadsheet_(), { source: 'auth', status: 'OK', message: 'Код надіслано: ' + email });
  } else {
    log_(getSpreadsheet_(), { source: 'auth', status: 'DENIED', message: 'Спроба входу з недозволеної адреси: ' + email });
  }
  return { sent: true };
}

/** Крок 2: перевірити код і видати токен сесії. */
function authVerify_(email, code) {
  email = normEmail_(email);
  code = String(code || '').replace(/\D/g, '');
  var cache = CacheService.getScriptCache(), key = 'code:' + sha_(email);
  var raw = cache.get(key);
  if (!raw) throw new Error('Код прострочений або не запитаний. Запросіть новий.');
  var rec = JSON.parse(raw);
  if (rec.a >= AUTH.MAX_ATTEMPTS) { cache.remove(key); throw new Error('Забагато спроб. Запросіть новий код.'); }
  if (rec.h !== sha_(email + ':' + code)) {
    rec.a++; cache.put(key, JSON.stringify(rec), AUTH.CODE_TTL_SEC);
    throw new Error('Невірний код');
  }
  cache.remove(key);
  var users = allowedUsers_();
  if (users[email] === undefined) throw new Error('Доступ не надано');

  return newSession_(email, users[email], 'код');
}

function newSession_(email, name, how) {
  var token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  var props = PropertiesService.getScriptProperties();
  props.setProperty('S_' + sha_(token), JSON.stringify({ e: email, x: Date.now() + AUTH.SESSION_DAYS * 864e5 }));
  cleanupSessions_(props);
  log_(getSpreadsheet_(), { source: 'auth', status: 'OK', message: 'Вхід (' + how + '): ' + email });
  return { token: token, email: email, name: name };
}

/** Вхід за паролем із аркуша _users (стовпець password). */
function authPassword_(email, password) {
  email = normEmail_(email);
  password = String(password || '');
  var cache = CacheService.getScriptCache(), key = 'pw:' + sha_(email);
  var n = parseInt(cache.get(key) || '0', 10);
  if (n >= AUTH.MAX_PASSWORD_TRIES) throw new Error('Забагато невдалих спроб. Спробуйте через 15 хвилин.');
  var u = usersTable_()[email];
  if (!u || !u.password || password !== u.password) {
    cache.put(key, String(n + 1), 900);
    log_(getSpreadsheet_(), { source: 'auth', status: 'DENIED', message: 'Невдала спроба входу: ' + email });
    throw new Error('Невірна пошта або пароль');
  }
  cache.remove(key);
  return newSession_(email, u.name, 'пароль');
}

/** Перевірка токена для кожного запиту даних. Повертає email. */
function authCheck_(token) {
  if (!token) throw new Error(AUTH_REQUIRED);
  var props = PropertiesService.getScriptProperties(), k = 'S_' + sha_(token);
  var raw = props.getProperty(k);
  if (!raw) throw new Error(AUTH_REQUIRED);
  var s = JSON.parse(raw);
  if (s.x < Date.now()) { props.deleteProperty(k); throw new Error(AUTH_REQUIRED); }
  if (allowedUsers_()[s.e] === undefined) { props.deleteProperty(k); throw new Error(AUTH_REQUIRED); }
  return s.e;
}

function authLogout_(token) {
  if (token) PropertiesService.getScriptProperties().deleteProperty('S_' + sha_(token));
  return { ok: true };
}

function cleanupSessions_(props) {
  var all = props.getProperties(), now = Date.now();
  Object.keys(all).forEach(function (k) {
    if (k.indexOf('S_') !== 0) return;
    try { if (JSON.parse(all[k]).x < now) props.deleteProperty(k); } catch (e) { props.deleteProperty(k); }
  });
}


// ===================== Api.js =====================
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
  clientItems: function (year, period, clientKey) { return apiClientItems(year, period, clientKey); },
  planfactMonths: function () { return apiPlanFactMonths(); },
  planfact: function (month) { return apiPlanFact(month); }
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
