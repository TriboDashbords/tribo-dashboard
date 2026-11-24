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

/** {email: {name, password}} для активних користувачів. */
function usersTable_() {
  var sh = getSpreadsheet_().getSheetByName(AUTH.USERS_SHEET), out = {};
  if (!sh || sh.getLastRow() < 2) return out;
  if (String(sh.getRange(1, 5).getValue()) === '') sh.getRange(1, 5).setValue('password');
  sh.getRange(2, 1, sh.getLastRow() - 1, 5).getDisplayValues().forEach(function (r) {
    var e = normEmail_(r[0]), active = String(r[2]).trim().toLowerCase();
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
