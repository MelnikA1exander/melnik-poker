// Приёмник результатов тренажёра «Профиль» → лист «Сессии» этой таблицы.
// Тренажёр шлёт POST с JSON сессии: {id, name, started, updated, summary, dec}.
// Повторная отправка той же сессии (тот же id) обновляет её строку, а не дублирует.

const SHEET = 'Сессии';
const HEAD = ['Получено', 'ID', 'Имя', 'Начало', 'Раздач', 'Тип', 'VPIP', 'PFR', 'Лимпы', 'Точность', 'Итог (JSON)', 'Решения (JSON)'];

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET);
  if (!sh) { sh = ss.insertSheet(SHEET); sh.appendRow(HEAD); sh.setFrozenRows(1); }
  return sh;
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const x = JSON.parse(e.postData.contents);
    if (!x || !x.id) return out_({ ok: false, error: 'no id' });
    const s = x.summary || {};
    const hands = s.hands || (x.dec || []).length;
    const pct = (v, n) => (v == null ? '' : v);
    const vpip = s.vpip != null ? s.vpip : (s.play != null && hands ? Math.round(s.play / hands * 100) : '');
    const pfr = s.pfr != null ? s.pfr : (s.raise != null && hands ? Math.round(s.raise / hands * 100) : '');
    const row = [
      new Date(), String(x.id), String(x.name || '').slice(0, 100),
      x.started ? new Date(x.started) : '', hands, s.type || '',
      vpip, pfr, s.limps != null ? s.limps : (x.dec || []).filter(d => d.k === 'open' && d.a === 'C').length, pct(s.acc),
      JSON.stringify(s).slice(0, 49000), JSON.stringify(x.dec || []).slice(0, 49000)
    ];
    const sh = sheet_();
    const ids = sh.getRange(1, 2, sh.getLastRow(), 1).getValues().map(r => String(r[0]));
    const at = ids.indexOf(String(x.id));
    if (at > 0) sh.getRange(at + 1, 1, 1, row.length).setValues([row]);
    else sh.appendRow(row);
    return out_({ ok: true });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() { sheet_(); return out_({ ok: true, sheet: SHEET }); }

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
