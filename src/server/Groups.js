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
