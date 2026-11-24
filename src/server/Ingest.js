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
