// Бэкенд тренажёра Melnik Poker на Google Таблице.
// Листы создаются сами: «Сессии» (результаты), «Раздачи профиля» (каждое решение тестов отдельной строкой, по времени), «Аккаунты» (логин, соль, хеш пароля), «Входы» (ключи входа).
// Все запросы — POST с JSON в теле (Content-Type: text/plain). Поле action выбирает действие:
//   register {login, password}       → {ok, token, login}
//   login    {login, password}       → {ok, token, login}
//   me       {token}                 → {ok, login}
//   logout   {token}                 → {ok}
//   stats    {token}                 → {ok, login, all, last100} — {sessions, hands, gto, acc, vpip, pfr, type}
//   preflop  {token, id, started, dec} → {ok} — личный журнал режима «Префлоп», лист «Префлоп».
//            Наружу не отдаётся никаким действием и в stats не входит; смотреть только в самой таблице.
//   без action — результат сессии {id, name, started, updated, summary, dec, token?};
//   повторная отправка той же сессии (тот же id) обновляет её строку.

const SHEET = 'Сессии', ACC = 'Аккаунты', TOK = 'Входы', PF = 'Префлоп', HANDS = 'Раздачи профиля';
const HEAD = ['Получено', 'ID', 'Имя', 'Начало', 'Раздач', 'Тип', 'VPIP', 'PFR', 'Лимпы', 'Точность', 'Итог (JSON)', 'Решения (JSON)', 'Аккаунт'];
const ACC_HEAD = ['Логин', 'Ключ', 'Соль', 'Хеш', 'Создан', 'Последний вход'];
const PF_HEAD = ['Получено', 'ID', 'Аккаунт', 'Начало', 'Решений', 'Верно %', 'Неверных', 'Решения (JSON)'];
const HANDS_HEAD = ['Время', 'Аккаунт', 'Тест (ID)', 'Раздача №', 'Позиция', 'Рука', 'Карты', 'Ситуация', 'К коллу', 'Действие', 'Сумма', 'Правильно', 'Оценка', 'Что было до тебя'];
const TOK_HEAD = ['Ключ входа', 'Аккаунт', 'Создан', 'Действует до'];
const TOKEN_DAYS = 90, HASH_ROUNDS = 2000, PROF_N = 50;

function sh_(name, head) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.appendRow(head); sh.setFrozenRows(1); }
  else if (sh.getLastColumn() < head.length) sh.getRange(1, 1, 1, head.length).setValues([head]);
  return sh;
}
const rows_ = sh => sh.getLastRow() < 2 ? [] : sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const x = JSON.parse(e.postData.contents);
    switch (x && x.action) {
      case 'register': return out_(register_(x));
      case 'login': return out_(login_(x));
      case 'me': { const a = auth_(x.token); return out_(a ? { ok: true, login: a.login } : { ok: false, error: 'auth' }); }
      case 'logout': return out_(logout_(x));
      case 'stats': return out_(stats_(x));
      case 'preflop': return out_(savePreflop_(x));
      default: return out_(saveSession_(x));
    }
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() { sh_(SHEET, HEAD); sh_(HANDS, HANDS_HEAD); sh_(PF, PF_HEAD); sh_(ACC, ACC_HEAD); sh_(TOK, TOK_HEAD); return out_({ ok: true, sheet: SHEET }); }

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// ---------- аккаунты ----------
const key_ = s => String(s || '').trim().toLowerCase();
function hash_(salt, pw) {
  let h = salt + '|' + pw;
  for (let i = 0; i < HASH_ROUNDS; i++) {
    h = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h, Utilities.Charset.UTF_8));
  }
  return h;
}
const rnd_ = () => Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');

function findAcc_(k) {
  const sh = sh_(ACC, ACC_HEAD), rs = rows_(sh);
  for (let i = 0; i < rs.length; i++) if (String(rs[i][1]) === k) return { sh, row: i + 2, login: String(rs[i][0]), salt: String(rs[i][2]), hash: String(rs[i][3]) };
  return null;
}
function issue_(acc) {
  const t = rnd_(), now = new Date();
  sh_(TOK, TOK_HEAD).appendRow([t, key_(acc.login), now, new Date(now.getTime() + TOKEN_DAYS * 864e5)]);
  return t;
}
// не больше 8 неудачных попыток входа на логин за 15 минут
function tooMany_(k) { return +(CacheService.getScriptCache().get('fail:' + k) || 0) >= 8; }
function fail_(k) { const c = CacheService.getScriptCache(), n = +(c.get('fail:' + k) || 0) + 1; c.put('fail:' + k, String(n), 900); }

function register_(x) {
  const login = String(x.login || '').trim(), k = key_(login), pw = String(x.password || '');
  // требования только к длине: логин 2–25 символов, пароль 4–25
  const ll = [...login].length, pl = [...pw].length;
  if (ll < 2 || ll > 25) return { ok: false, error: 'login_format' };
  if (pl < 4 || pl > 25) return { ok: false, error: 'password_format' };
  if (findAcc_(k)) return { ok: false, error: 'login_taken' };
  const salt = rnd_(), now = new Date();
  sh_(ACC, ACC_HEAD).appendRow([login, k, salt, hash_(salt, pw), now, now]);
  const acc = { login };
  return { ok: true, token: issue_(acc), login };
}

function login_(x) {
  const k = key_(x.login), pw = String(x.password || '');
  if (tooMany_(k)) return { ok: false, error: 'too_many' };
  const acc = findAcc_(k);
  if (!acc || hash_(acc.salt, pw) !== acc.hash) { fail_(k); return { ok: false, error: 'bad_credentials' }; }
  acc.sh.getRange(acc.row, 6).setValue(new Date());
  return { ok: true, token: issue_(acc), login: acc.login };
}

function auth_(token) {
  if (!token) return null;
  const rs = rows_(sh_(TOK, TOK_HEAD)), now = Date.now();
  for (const r of rs) if (String(r[0]) === String(token) && new Date(r[3]).getTime() > now) {
    const acc = findAcc_(String(r[1]));
    return acc ? { login: acc.login, key: key_(acc.login) } : null;
  }
  return null;
}

function logout_(x) {
  const sh = sh_(TOK, TOK_HEAD), rs = rows_(sh);
  for (let i = rs.length - 1; i >= 0; i--) if (String(rs[i][0]) === String(x.token)) sh.deleteRow(i + 2);
  return { ok: true };
}

// ---------- сессии ----------
function saveSession_(x) {
  if (!x || !x.id) return { ok: false, error: 'no id' };
  if (/^pf/.test(String(x.id))) return { ok: false, error: 'preflop' }; // журнал «Префлопа» в «Сессии» не пишем
  // пока записываем только завершённые тесты «Профиля»
  if ((x.summary && x.summary.hands || (x.dec || []).length) < PROF_N) return { ok: false, error: 'unfinished' };
  const a = auth_(x.token);
  const s = x.summary || {}, dec = x.dec || [];
  const hands = s.hands || dec.length;
  const vpip = s.vpip != null ? s.vpip : (s.play != null && hands ? Math.round(s.play / hands * 100) : '');
  const pfr = s.pfr != null ? s.pfr : (s.raise != null && hands ? Math.round(s.raise / hands * 100) : '');
  const limps = s.limps != null ? s.limps : dec.filter(d => d.k === 'open' && d.a === 'C').length;
  const row = [
    new Date(), String(x.id), String(a ? a.login : (x.name || '')).slice(0, 100),
    x.started ? new Date(x.started) : '', hands, s.type || '',
    vpip, pfr, limps, s.acc == null ? '' : s.acc,
    JSON.stringify(s).slice(0, 49000), JSON.stringify(dec).slice(0, 49000), a ? a.key : ''
  ];
  const sh = sh_(SHEET, HEAD);
  const ids = sh.getRange(1, 2, sh.getLastRow(), 1).getValues().map(r => String(r[0]));
  const at = ids.indexOf(String(x.id));
  if (at > 0) sh.getRange(at + 1, 1, 1, row.length).setValues([row]);
  else { sh.appendRow(row); addHands_(String(x.id), a ? a.login : String(x.name || ''), dec); }
  return { ok: true };
}

// ---------- раздачи профиля: одна строка на решение, в порядке времени ----------
const SIT_ = { open: 'Первым в банк', iso: 'Против лимпа', bbiso: 'ББ против лимпа', vsopen: 'Против рейза', vs3b: 'Против 3-бета',
  cold3b: 'Против рейза и 3-бета', vs4b: 'Против 4-бета', cold4b: 'Против 4-бета, холодно', vsjam: 'Против олл-ина' };
const RNAME_ = { iso: 'Рейз', open: 'Рейз', bbiso: 'Рейз', vsopen: '3-бет', vs3b: '4-бет', cold3b: '4-бет', vs4b: '5-бет', cold4b: '5-бет', vsjam: 'Олл-ин' };
function actName_(a, d) {
  if (a === 'F') return 'Фолд'; if (a === 'X') return 'Чек';
  if (a === 'C') return d.k === 'open' ? 'Лимп' : 'Колл';
  return RNAME_[d.k] || 'Рейз';
}
function handRows_(id, login, dec) {
  let n = 0, prev = null;
  return dec.map(d => {
    // номер раздачи: из клиента, а для старых тестов — по смене карт или позиции
    const key = (d.c || d.h) + '|' + d.pos;
    if (d.n) n = d.n; else if (key !== prev) n++;
    prev = key;
    return [d.t ? new Date(d.t) : '', login, id, n, d.pos || '', d.h || '', d.c || '', SIT_[d.k] || d.k || '',
      d.toCall || 0, actName_(d.a, d), d.a === 'R' ? d.amt || '' : '',
      Array.isArray(d.ok) ? d.ok.map(a => actName_(a, d)).join(' / ') : '',
      d.gr === 2 ? 'верно' : d.gr === 1 ? 'неточный размер' : d.gr === 0 ? 'ошибка' : '', String(d.ctx || '').slice(0, 300)];
  });
}
function addHands_(id, login, dec) {
  const rows = handRows_(id, login, dec || []);
  if (!rows.length) return;
  const sh = sh_(HANDS, HANDS_HEAD);
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, HANDS_HEAD.length).setValues(rows);
}
// Запусти один раз вручную в редакторе (выбрать rebuildHands → Выполнить): пересобирает «Раздачи профиля»
// из всех тестов в «Сессиях», включая сохранённые до появления этого листа. Повторный запуск безопасен.
function rebuildHands() {
  const all = [];
  rows_(sh_(SHEET, HEAD)).forEach(r => {
    if (/^pf/.test(String(r[1]))) return;
    let dec = []; try { dec = JSON.parse(r[11] || '[]'); } catch (e) { return; }
    handRows_(String(r[1]), String(r[2]), dec).forEach(x => all.push(x));
  });
  all.sort((p, q) => (p[0] ? +p[0] : 0) - (q[0] ? +q[0] : 0));
  const sh = sh_(HANDS, HANDS_HEAD);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).clearContent();
  if (all.length) sh.getRange(2, 1, all.length, HANDS_HEAD.length).setValues(all);
  return all.length;
}

// журнал «Префлопа»: одна строка на серию раздач, повторная отправка того же id обновляет строку
function savePreflop_(x) {
  const a = auth_(x.token);
  if (!a) return { ok: false, error: 'auth' };
  if (!x.id || !/^pf/.test(String(x.id))) return { ok: false, error: 'no id' };
  const dec = Array.isArray(x.dec) ? x.dec.slice(0, 300) : [];
  if (!dec.length) return { ok: true };
  const good = dec.filter(d => d && d.g === 2).length, bad = dec.filter(d => d && d.g === 0).length;
  const row = [new Date(), String(x.id), a.login, x.started ? new Date(x.started) : '', dec.length,
    Math.round(good / dec.length * 100), bad, JSON.stringify(dec).slice(0, 49000)];
  const sh = sh_(PF, PF_HEAD);
  const ids = sh.getRange(1, 2, sh.getLastRow(), 1).getValues().map(r => String(r[0]));
  const at = ids.indexOf(String(x.id));
  if (at > 0) {
    if (String(sh.getRange(at + 1, 3).getValue()) !== a.login) return { ok: false, error: 'owner' };
    sh.getRange(at + 1, 1, 1, row.length).setValues([row]);
  } else sh.appendRow(row);
  return { ok: true };
}

// общая статистика игрока без конкретных раздач: за всё время и за последние 100 раздач
function stats_(x) {
  const a = auth_(x.token);
  if (!a) return { ok: false, error: 'auth' };
  const num = v => (v === '' || v == null || isNaN(+v)) ? null : +v;
  const list = rows_(sh_(SHEET, HEAD))
    .filter(r => String(r[12]) === a.key && +r[4] >= PROF_N && !/^pf/.test(String(r[1]))) // pf… — журнал «Префлопа», попавший сюда со старой версией скрипта
    .map(r => {
      let sm = {}; try { sm = JSON.parse(r[10] || '{}'); } catch (e) {}
      return { date: r[3] ? new Date(r[3]).getTime() : new Date(r[0]).getTime(), hands: +r[4], type: String(r[5]),
        gto: num(sm.gto), acc: num(r[9]), vpip: num(r[6]), pfr: num(r[7]) };
    })
    .sort((p, q) => p.date - q.date);
  const recent = []; let n = 0;
  for (let i = list.length - 1; i >= 0 && n < 100; i--) { recent.unshift(list[i]); n += list[i].hands; }
  return { ok: true, login: a.login, all: total_(list), last100: total_(recent) };
}

function total_(list) {
  const avg = f => { let w = 0, s = 0; for (const r of list) if (r[f] != null) { s += r[f] * r.hands; w += r.hands; } return w ? Math.round(s / w) : null; };
  // тип: самый частый среди тестов периода, при равенстве — более свежий
  const cnt = {}; let type = '', best = 0;
  list.forEach((r, i) => { if (!r.type) return; cnt[r.type] = (cnt[r.type] || 0) + 1 + i * 1e-6; if (cnt[r.type] >= best) { best = cnt[r.type]; type = r.type; } });
  return { sessions: list.length, hands: list.reduce((t, r) => t + r.hands, 0), gto: avg('gto'), acc: avg('acc'), vpip: avg('vpip'), pfr: avg('pfr'), type };
}
