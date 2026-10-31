// Бэкенд тренажёра Melnik Poker на Google Таблице.
// Листы создаются сами: «Сессии» (результаты), «Раздачи профиля» (каждое решение тестов отдельной строкой, по времени), «Аккаунты» (логин, соль, хеш пароля), «Входы» (ключи входа),
// «Последние 300» (по каждому аккаунту его последние 300 префлоп-решений из «Профиля» и «Префлопа»).
// Все запросы — POST с JSON в теле (Content-Type: text/plain). Поле action выбирает действие:
//   register {login, password}       → {ok, token, login}
//   login    {login, password}       → {ok, token, login}
//   me       {token}                 → {ok, login}
//   logout   {token}                 → {ok}
//   stats    {token}                 → {ok, login, all, last300, last100} — {sessions, hands, gto, acc, vpip, pfr, type}
//   preflop  {token, id, started, dec} → {ok} — личный журнал режима «Префлоп», лист «Префлоп».
//            Наружу не отдаётся никаким действием и в stats не входит; смотреть только в самой таблице.
//   history  {token}                 → {ok, dec} — последние ≤300 префлоп-решений игрока (тесты «Профиля» + журнал «Префлопа»)
//            из листа «Последние 300»; тренажёр берёт их как стартовую точку для оценки профиля.
//   без action — результат сессии {id, name, started, updated, summary, dec, token?};
//   повторная отправка той же сессии (тот же id) обновляет её строку.

const SHEET = 'Сессии', ACC = 'Аккаунты', TOK = 'Входы', PF = 'Префлоп', HANDS = 'Раздачи профиля', RECENT = 'Последние 300';
const HEAD = ['Получено', 'ID', 'Имя', 'Начало', 'Раздач', 'Тип', 'VPIP', 'PFR', 'Лимпы', 'Точность', 'Итог (JSON)', 'Решения (JSON)', 'Аккаунт'];
const ACC_HEAD = ['Логин', 'Ключ', 'Соль', 'Хеш', 'Создан', 'Последний вход'];
const PF_HEAD = ['Получено', 'ID', 'Аккаунт', 'Начало', 'Решений', 'Верно %', 'Неверных', 'Решения (JSON)'];
// «Раздачи профиля» — указатель: 3 ячейки на решение. Ключ «аккаунт/тест/№ раздачи» ведёт к полной записи теста в «Сессиях».
const HANDS_HEAD = ['Время', 'Ключ', 'Раздача'];
const RECENT_HEAD = ['Время', 'Аккаунт', 'Источник', 'Тест / серия', 'Позиция', 'Рука', 'Карты', 'Ситуация', 'К коллу', 'Действие', 'Сумма', 'Правильно', 'Оценка', 'Данные (JSON)'];
const RECENT_N = 300;
const TOK_HEAD = ['Ключ входа', 'Аккаунт', 'Создан', 'Действует до'];
const TOKEN_DAYS = 90, HASH_ROUNDS = 2000, PROF_N = 50;

// Лимит Google Таблицы — 10 млн ячеек, и считаются все ячейки сетки, даже пустые.
// Поэтому новый лист сразу обрезается до ширины заголовка и двух строк, а строки добавляются по мере записи.
function sh_(name, head) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, head.length).setValues([head]); sh.setFrozenRows(1);
    fit_(sh, head.length);
  } else if (sh.getLastColumn() < head.length) {
    if (sh.getMaxColumns() < head.length) sh.insertColumnsAfter(sh.getMaxColumns(), head.length - sh.getMaxColumns());
    sh.getRange(1, 1, 1, head.length).setValues([head]);
  }
  return sh;
}
// убирает пустые строки и столбцы за пределами данных (оставляет минимум 2 строки)
function fit_(sh, width) {
  const w = Math.max(width || 1, sh.getLastColumn()), mc = sh.getMaxColumns();
  if (mc > w) sh.deleteColumns(w + 1, mc - w);
  const h = Math.max(2, sh.getLastRow()), mr = sh.getMaxRows();
  if (mr > h) sh.deleteRows(h + 1, mr - h);
}
// дописывает строки в конец листа, при нехватке места сначала добавляет строки в сетку
function append_(sh, rows) {
  if (!rows.length) return;
  const at = sh.getLastRow() + 1, need = at + rows.length - 1 - sh.getMaxRows();
  if (need > 0) sh.insertRowsAfter(sh.getMaxRows(), need);
  sh.getRange(at, 1, rows.length, rows[0].length).setValues(rows);
}
// записывает блок строк начиная со 2-й, расширяя сетку при нехватке места
function put_(sh, rows, width) {
  const need = 1 + rows.length - sh.getMaxRows();
  if (need > 0) sh.insertRowsAfter(sh.getMaxRows(), need);
  if (rows.length) sh.getRange(2, 1, rows.length, width).setValues(rows);
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
      case 'history': return out_(history_(x));
      default: return out_(saveSession_(x));
    }
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() { sh_(SHEET, HEAD); sh_(HANDS, HANDS_HEAD); sh_(PF, PF_HEAD); sh_(RECENT, RECENT_HEAD); sh_(ACC, ACC_HEAD); sh_(TOK, TOK_HEAD); return out_({ ok: true, sheet: SHEET }); }

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
  append_(sh_(TOK, TOK_HEAD), [[t, key_(acc.login), now, new Date(now.getTime() + TOKEN_DAYS * 864e5)]]);
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
  append_(sh_(ACC, ACC_HEAD), [[login, k, salt, hash_(salt, pw), now, now]]);
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
  else { append_(sh, [row]); addHands_(String(x.id), a ? a.login : String(x.name || ''), dec); }
  if (a) updRecent_(a, String(x.id), 'Профиль', dec);
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
// Одна раздача — одна строка из 3 ячеек: время, ключ «аккаунт/тест/№» и сама раздача текстом, например
// «BTN AhKd (AKo) · Против рейза: UTG+1 14 · к коллу 14 → 3-бет 44 · GTO: 3-бет / Колл · ✓ · до: UTG+1 рейз 14 · UTG+2 колл».
// ✓ — верно, ≈ — неточный размер, ✗ — ошибка. Полная запись решения со всеми полями — в «Сессиях», колонка «Решения (JSON)».
function handRows_(id, login, dec) {
  let n = 0, prev = null;
  return dec.map(d => {
    // номер раздачи: из клиента, а для старых тестов — по смене карт или позиции
    const key = (d.c || d.h) + '|' + d.pos;
    if (d.n) n = d.n; else if (key !== prev) n++;
    prev = key;
    const parts = [[d.pos, d.c, d.h && d.c ? '(' + d.h + ')' : d.h].filter(Boolean).join(' ')];
    parts.push((SIT_[d.k] || d.k || '') + (d.opos ? ': ' + d.opos + (d.open ? ' ' + d.open : '') : '') + (d.lim ? ', лимперов ' + d.lim : ''));
    parts.push((d.toCall ? 'к коллу ' + d.toCall + ' → ' : '→ ') + actName_(d.a, d) + (d.a === 'R' && d.amt ? ' ' + d.amt : ''));
    if (Array.isArray(d.ok) && d.ok.length) parts.push('GTO: ' + d.ok.map(a => actName_(a, d)).join(' / '));
    parts.push(d.gr === 2 ? '✓' : d.gr === 1 ? '≈' : d.gr === 0 ? '✗' : '?');
    if (d.ctx) parts.push('до: ' + String(d.ctx).slice(0, 200));
    return [d.t ? new Date(d.t) : '', login + '/' + id + '/' + n, parts.join(' · ')];
  });
}
function addHands_(id, login, dec) {
  append_(sh_(HANDS, HANDS_HEAD), handRows_(id, login, dec || []));
}
// Пересобирает «Раздачи профиля» из всех тестов в «Сессиях». Повторный запуск безопасен.
function rebuildHands() {
  const all = [];
  rows_(sh_(SHEET, HEAD)).forEach(r => {
    if (/^pf/.test(String(r[1]))) return;
    let dec = []; try { dec = JSON.parse(r[11] || '[]'); } catch (e) { return; }
    handRows_(String(r[1]), String(r[2]), dec).forEach(x => all.push(x));
  });
  all.sort((p, q) => (p[0] ? +p[0] : 0) - (q[0] ? +q[0] : 0));
  const ss = SpreadsheetApp.getActiveSpreadsheet(), old = ss.getSheetByName(HANDS);
  if (old) ss.deleteSheet(old); // старый формат был шире — проще создать лист заново
  append_(sh_(HANDS, HANDS_HEAD), all);
  return all.length;
}

// Запусти один раз вручную (compactStorage → Выполнить): переводит «Раздачи профиля» в формат 3 ячеек,
// обрезает пустые строки и столбцы во всех листах, удаляет пустой «Sheet1»/«Лист1» и просроченные ключи входа.
// Возвращает, сколько ячеек занято из 10 млн. Повторный запуск безопасен.
function compactStorage() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  rebuildHands();
  ['Sheet1', 'Лист1'].forEach(n => { const x = ss.getSheetByName(n); if (x && x.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(x); });
  const tok = sh_(TOK, TOK_HEAD), now = Date.now(), tr = rows_(tok);
  for (let i = tr.length - 1; i >= 0; i--) if (new Date(tr[i][3]).getTime() < now) tok.deleteRow(i + 2);
  ss.getSheets().forEach(x => fit_(x, x.getLastColumn()));
  const used = cellsUsed();
  Logger.log('Занято ячеек: ' + used + ' из 10 000 000');
  return used;
}
// сколько ячеек сетки занято во всей таблице (лимит Google — 10 000 000)
function cellsUsed() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets().reduce((t, x) => t + x.getMaxRows() * x.getMaxColumns(), 0);
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
  } else append_(sh, [row]);
  updRecent_(a, String(x.id), 'Префлоп', dec);
  return { ok: true };
}

// ---------- последние 300 решений игрока ----------
// Строка на решение; в последнем столбце — данные для тренажёра (позиция, рука, ситуация, действие, размер, правильные действия).
function recentRow_(login, src, id, d) {
  const g = d.gr != null ? d.gr : d.g;
  const data = { t: d.t || 0, src, pos: d.pos || '', h: d.h || '', k: d.k || '', a: d.a || '', amt: d.amt || 0,
    ok: Array.isArray(d.ok) ? d.ok : [], toCall: d.toCall || 0, open: d.open || null, lim: d.lim || 0, cur: d.cur || null, N: d.N || 9 };
  return [d.t ? new Date(d.t) : '', login, src, id, d.pos || '', d.h || '', d.c || '', SIT_[d.k] || d.k || '', d.toCall || 0, actName_(d.a, d),
    d.a === 'R' ? d.amt || '' : '', Array.isArray(d.ok) ? d.ok.map(a => actName_(a, d)).join(' / ') : '',
    g === 2 ? 'верно' : g === 1 ? 'неточный размер' : g === 0 ? 'ошибка' : '', JSON.stringify(data)];
}
const recentDec_ = d => d && d.h && d.k && d.a; // только префлоп-решения с рукой и ситуацией
// все решения аккаунта из «Сессий» и «Префлопа»: [источник, id, решение]
function allDec_(a) {
  const out = [];
  rows_(sh_(SHEET, HEAD)).forEach(r => {
    if (String(r[12]) !== a.key || /^pf/.test(String(r[1]))) return;
    let d = []; try { d = JSON.parse(r[11] || '[]'); } catch (e) { return; }
    d.forEach(x => out.push(['Профиль', String(r[1]), x]));
  });
  rows_(sh_(PF, PF_HEAD)).forEach(r => {
    if (key_(r[2]) !== a.key) return;
    let d = []; try { d = JSON.parse(r[7] || '[]'); } catch (e) { return; }
    d.forEach(x => out.push(['Префлоп', String(r[1]), x]));
  });
  return out;
}
// обновляет строки аккаунта: заменяет решения теста/серии id новыми и оставляет последние 300;
// если строк аккаунта ещё нет — собирает их из всей истории
function updRecent_(a, id, src, dec) {
  const sh = sh_(RECENT, RECENT_HEAD), W = RECENT_HEAD.length, all = rows_(sh).map(r => r.slice(0, W));
  const others = all.filter(r => key_(r[1]) !== a.key);
  let mine = all.filter(r => key_(r[1]) === a.key);
  if (!mine.length) mine = allDec_(a).filter(x => recentDec_(x[2])).map(x => recentRow_(a.login, x[0], x[1], x[2]));
  else {
    mine = mine.filter(r => String(r[3]) !== String(id));
    (dec || []).filter(recentDec_).forEach(d => mine.push(recentRow_(a.login, src, String(id), d)));
  }
  const tm = r => r[0] ? new Date(r[0]).getTime() : 0;
  mine.sort((p, q) => tm(p) - tm(q));
  const out = others.concat(mine.slice(-RECENT_N));
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).clearContent();
  put_(sh, out, W);
}
function history_(x) {
  const a = auth_(x.token);
  if (!a) return { ok: false, error: 'auth' };
  const mineRows = () => rows_(sh_(RECENT, RECENT_HEAD)).filter(r => key_(r[1]) === a.key);
  let mine = mineRows();
  if (!mine.length) { updRecent_(a, '', '', []); mine = mineRows(); }
  const dec = [];
  mine.forEach(r => { try { dec.push(JSON.parse(r[13])); } catch (e) {} });
  return { ok: true, dec };
}
// Запусти один раз вручную (rebuildRecent → Выполнить): пересобирает «Последние 300» для всех аккаунтов из всей истории.
function rebuildRecent() {
  const sh = sh_(RECENT, RECENT_HEAD);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).clearContent();
  rows_(sh_(ACC, ACC_HEAD)).forEach(r => updRecent_({ login: String(r[0]), key: String(r[1]) }, '', '', []));
}

// общая статистика игрока без конкретных раздач: за всё время и за последние 100 раздач
function stats_(x) {
  const a = auth_(x.token);
  if (!a) return { ok: false, error: 'auth' };
  const num = v => (v === '' || v == null || isNaN(+v)) ? null : +v;
  // колонку «Решения (JSON)» (самую тяжёлую) не читаем: берём A:K и аккаунт из M
  const sh = sh_(SHEET, HEAD), cnt = sh.getLastRow() - 1;
  const base = cnt > 0 ? sh.getRange(2, 1, cnt, 11).getValues() : [], accs = cnt > 0 ? sh.getRange(2, 13, cnt, 1).getValues() : [];
  const list = base.map((r, i) => r.concat(['', accs[i][0]]))
    .filter(r => String(r[12]) === a.key && +r[4] >= PROF_N && !/^pf/.test(String(r[1]))) // pf… — журнал «Префлопа», попавший сюда со старой версией скрипта
    .map(r => {
      let sm = {}; try { sm = JSON.parse(r[10] || '{}'); } catch (e) {}
      return { date: r[3] ? new Date(r[3]).getTime() : new Date(r[0]).getTime(), hands: +r[4], type: String(r[5]),
        gto: num(sm.gto), acc: num(r[9]), vpip: num(r[6]), pfr: num(r[7]) };
    })
    .sort((p, q) => p.date - q.date);
  const lastN = N => { const out = []; let n = 0; for (let i = list.length - 1; i >= 0 && n < N; i--) { out.unshift(list[i]); n += list[i].hands; } return out; };
  return { ok: true, login: a.login, all: total_(list), last300: total_(lastN(300)), last100: total_(lastN(100)) };
}

function total_(list) {
  const avg = f => { let w = 0, s = 0; for (const r of list) if (r[f] != null) { s += r[f] * r.hands; w += r.hands; } return w ? Math.round(s / w) : null; };
  // тип: самый частый среди тестов периода, при равенстве — более свежий
  const cnt = {}; let type = '', best = 0;
  list.forEach((r, i) => { if (!r.type) return; cnt[r.type] = (cnt[r.type] || 0) + 1 + i * 1e-6; if (cnt[r.type] >= best) { best = cnt[r.type]; type = r.type; } });
  return { sessions: list.length, hands: list.reduce((t, r) => t + r.hands, 0), gto: avg('gto'), acc: avg('acc'), vpip: avg('vpip'), pfr: avg('pfr'), type };
}
