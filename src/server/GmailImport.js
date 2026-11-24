/**
 * GmailImport.js — автоматичне завантаження звітів, які 1С надсилає поштою
 * («Розсилки звітів» у 1С → лист із вкладенням .xlsx на пошту власника дашборду).
 *
 * Викликається з processInbox (тригер кожні 15 хв). Беруться лише листи:
 *   - з вкладенням .xlsx,
 *   - від відправників з аркуша "_senders" (email | note) — захист від чужих листів,
 *   - ще не позначені міткою "tribo-done" чи "tribo-error" (мітки ставить скрипт).
 * Тип звіту (Продажи / План-факт) визначається автоматично, як і для папки Inbox.
 * Потрібен дозвіл Gmail у манифесті (loader/appsscript.json) — без нього крок тихо пропускається.
 */

var GMAIL = {
  SENDERS_SHEET: '_senders',
  LABEL_DONE: 'tribo-done',     // латиницею — так надійніше працює пошук Gmail
  LABEL_ERROR: 'tribo-error',
  QUERY: 'has:attachment filename:xlsx newer_than:7d -label:tribo-done -label:tribo-error',
  MAX_THREADS: 10
};

function allowedSenders_() {
  var sh = getOrCreateSheet_(getSpreadsheet_(), GMAIL.SENDERS_SHEET, ['email', 'note']);
  var out = {};
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 1).getDisplayValues().forEach(function (r) {
    var e = normEmail_(r[0]); if (e) out[e] = true;
  });
  return out;
}

function senderEmail_(from) {
  var m = /<([^>]+)>/.exec(String(from || ''));
  return normEmail_(m ? m[1] : from);
}

/** Повертає кількість оброблених вкладень. started — час початку (для ліміту 6 хв). */
function processGmail_(started) {
  var GA;
  try { GA = GmailApp; GA.getUserLabelByName(GMAIL.LABEL_DONE); } catch (e) {
    console.log('Gmail недоступний (немає дозволу в манифесті) — пропускаю: ' + e);
    return 0;
  }
  var ss = getSpreadsheet_();
  var done = GA.getUserLabelByName(GMAIL.LABEL_DONE) || GA.createLabel(GMAIL.LABEL_DONE);
  var bad = GA.getUserLabelByName(GMAIL.LABEL_ERROR) || GA.createLabel(GMAIL.LABEL_ERROR);
  var senders = allowedSenders_();
  var threads = GA.search(GMAIL.QUERY, 0, GMAIL.MAX_THREADS);
  var count = 0;
  threads.forEach(function (th) {
    if (Date.now() - started > 4 * 60 * 1000) return;
    var ok = true, any = false;
    th.getMessages().forEach(function (msg) {
      var from = senderEmail_(msg.getFrom());
      var atts = msg.getAttachments({ includeInlineImages: false }).filter(function (a) { return /\.xlsx$/i.test(a.getName()); });
      if (!atts.length) return;
      if (!senders[from]) {
        log_(ss, { source: 'gmail', status: 'SKIP', message: 'Лист від ' + from + ' («' + msg.getSubject() + '») пропущено: адреси немає в аркуші _senders' });
        ok = false; return;
      }
      atts.forEach(function (a) {
        any = true;
        try {
          var r = importSalesXlsxBlob_(a.copyBlob(), 'пошта: ' + a.getName());
          if (!r.ok) ok = false; else count++;
        } catch (err) {
          ok = false;
          log_(ss, { source: 'gmail', status: 'ERROR', message: a.getName() + ': ' + err });
        }
      });
    });
    th.addLabel(ok && any ? done : bad);
    if (ok && any) th.markRead();
  });
  return count;
}
