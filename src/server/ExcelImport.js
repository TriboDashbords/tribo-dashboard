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
    if (!/\.xlsx$/i.test(name)) continue; // сторонні файли (не Excel) не чіпаємо і не засмічуємо журнал
    try {
      var result = importSalesXlsxBlob_(f.getBlob(), name);
      f.moveTo(result.ok ? done : bad);
    } catch (err) {
      log_(ss, { source: 'excel', dataset: 'sales', status: 'ERROR', message: name + ': ' + err });
      f.moveTo(bad);
    }
  }

  // прибирання: оброблені файли зберігаються добу, з помилками — 7 днів, далі видаляються остаточно
  try {
    cleanupFolder_(done, CONFIG.KEEP_DONE_HOURS);
    cleanupFolder_(bad, CONFIG.KEEP_ERROR_HOURS);
  } catch (err) { log_(ss, { source: 'cleanup', status: 'ERROR', message: String(err) }); }

  // звіти, що 1С надсилає поштою (Розсилки звітів)
  try { processGmail_(started); } catch (err) { log_(ss, { source: 'gmail', status: 'ERROR', message: String(err) }); }
}

/** Остаточно видаляє файли папки, старші за maxAgeHours (не в Кошик, щоб не займати місце). */
function cleanupFolder_(folder, maxAgeHours) {
  var limit = Date.now() - maxAgeHours * 3600 * 1000, n = 0, started = Date.now();
  var it = folder.getFiles();
  while (it.hasNext() && Date.now() - started < 60 * 1000) {
    var f = it.next();
    if (f.getLastUpdated().getTime() > limit) continue;
    deleteFileForever_(f.getId()) || f.setTrashed(true);
    n++;
  }
  if (n) console.log('Видалено старих файлів: ' + n + ' з ' + folder.getName());
  return n;
}

function deleteFileForever_(id) {
  try {
    var r = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + id, {
      method: 'delete', muteHttpExceptions: true,
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }
    });
    return r.getResponseCode() === 204;
  } catch (e) { return false; }
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
