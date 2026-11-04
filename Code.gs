// Google Apps Script: приймає результати зі сторінки й записує їх у Google Таблицю.
// Скрипт створюйте з самої таблиці: Розширення → Apps Script.
const SHEET_NAME = 'Результати';

function norm(s) { return String(s).toLowerCase().replace(/\s+/g, ' ').trim(); }
function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const d = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    if (sh.getLastRow() === 0) sh.appendRow(['Час', 'Прізвище та ім’я', 'Клас', 'Частина', 'Бали', 'Максимум', 'Відповіді (номери варіантів)']);
    const key = norm(d.name) + '|' + norm(d.cls) + '|' + norm(d.part);
    const rows = sh.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (norm(rows[i][1]) + '|' + norm(rows[i][2]) + '|' + norm(rows[i][3]) === key) return out({ ok: false, error: 'duplicate' });
    }
    sh.appendRow([new Date(), d.name, d.cls, d.part, d.score, d.max, d.details]);
    return out({ ok: true });
  } catch (err) {
    return out({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
