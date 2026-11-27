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
  KEEP_DONE_HOURS: 24,      // скільки зберігати оброблені звіти в Inbox/Оброблено
  KEEP_ERROR_HOURS: 168,    // скільки зберігати файли з помилками в Inbox/Помилки

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
