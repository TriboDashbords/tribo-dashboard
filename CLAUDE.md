# ТРІБО · Дашборд продажів — правила для Claude

## Як це працює
- Google Таблиця «Трібо Дані» (акаунт tribodashbords@gmail.com) — сховище даних. Її Apps Script містить ЛИШЕ `loader/Code.gs`
  (вставлено вручну один раз). Він при кожному виклику бере код звідси: `dist/server.js` і `dist/index.html`
  з гілки `main` через raw.githubusercontent.com (кеш 2 хв). Тобто **push у main = зміна на робочому дашборді за ~2–5 хв**.
- Одна адреса веб-застосунку (Execute as: Me, Access: Anyone):
  - GET → сторінка дашборду (вхід за кодом із пошти, `src/server/Auth.js`; дозволені адреси — аркуш `_users`);
  - GET `?health=1` → перевірка; POST → прийом JSON з 1С (`src/server/Ingest.js`, токен у Script Properties `TOKEN`).
- Сторінка викликає сервер лише через `google.script.run.api(name, args)`; список дій — `PUBLIC_API`/`PRIVATE_API` у `src/server/Api.js`.
  Нова серверна функція для сторінки = новий запис у `PRIVATE_API` (заглушку в loader міняти не треба).
- Службова задача: власник натискає `run` у редакторі → виконується `RUN_TASK` з `Api.js`. Змінюй її під потрібну разову дію.
- Заглушки в loader: doGet, doPost, api, processInbox (тригер кожні 15 хв), setup, showToken, run, onOpen. Нових додати не можна без участі користувача.
- Зміни в `loader/` або в манифесті (`loader/appsscript.json`, нові OAuth-скоупи) потребують ручної вставки користувачем + нової версії розгортання. Уникай.

## Робочий процес
1. Редагуй `src/server/*.js` і `src/web/index.html`.
2. `node build.js && node tests/run.js` — усі перевірки мають пройти (імітація Apps Script у tests/run.js).
3. Коміть і src, і dist; push у `main`.

## Обмеження
- Репозиторій ПУБЛІЧНИЙ: жодних реальних даних продажів, сум, назв клієнтів, токенів, пошт користувачів. Тестові дані — вигадані (`TestData.js`).
- Apps Script V8, без npm-бібліотек; ліміт 6 хв на виклик; CacheService ≤100 КБ на значення (є chunk-хелпери).
- `safeDeleteRows_` замість `deleteRows`, коли можна видалити всі рядки даних; перед записом за межі аркуша — `insertRowsAfter`.
- Інтерфейс — українською, mobile-first. Суми в тис. грн, 1 знак, роздільник тисяч.
- Специфікація v1 і ТЗ для 1С — у документах користувача (claude.ai); формат JSON — розділ 4/5.1 (див. `src/server/Config.js`).
