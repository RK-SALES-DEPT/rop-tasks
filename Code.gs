/**
 * Задачи РОП — Telegram-бот + API для Mini App.
 * Хранилище: эта Google Таблица (листы «Задачи» и «Ключевые слова»).
 *
 * Свойства скрипта (Настройки проекта → Свойства скрипта):
 *   BOT_TOKEN   — токен бота из @BotFather (вводится вручную, нигде больше не хранится)
 *   WEBAPP_URL  — URL развёртывания этого скрипта (…/exec)
 *   MINIAPP_URL — адрес Mini App (GitHub Pages)
 *   OWNER_ID    — заполняется сам при первом /start
 */

const TZ = 'Europe/Moscow';
const SHEET_TASKS = 'Задачи';
const SHEET_KEYS = 'Ключевые слова';
const HEAD = ['id', 'Создана', 'Задача', 'Направление', 'Кто', 'Срок', 'Статус', 'Выполнена', 'Источник'];
const COL = { id: 1, created: 2, title: 3, dir: 4, person: 5, due: 6, status: 7, doneAt: 8, source: 9 };
const STATUS_RU = { inbox: 'Входящие', active: 'В работе', waiting: 'Жду', idea: 'Идея', done: 'Готово' };
const DIR_CODES = { t: 'ТЛМ', o: 'Исходящие', i: 'Входящие' };
const WD_SHORT = ['вс', 'пн', 'вт', 'ср', 'чт', 'пт', 'сб'];
const MONTHS = ['января', 'февраля', 'марта', 'апреля', 'мая', 'июня', 'июля', 'августа', 'сентября', 'октября', 'ноября', 'декабря'];

/* ───────────── Настройка ───────────── */

function setup() {
  const ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetTimeZone(TZ);

  let sh = ss.getSheetByName(SHEET_TASKS) || ss.insertSheet(SHEET_TASKS, 0);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEAD);
    sh.setFrozenRows(1);
    sh.getRange('1:1').setFontWeight('bold');
    sh.getRange('A:I').setNumberFormat('@');
    sh.setColumnWidth(COL.title, 360);
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(Object.keys(STATUS_RU), true).setAllowInvalid(false).build();
    sh.getRange('G2:G').setDataValidation(rule);
  }

  let k = ss.getSheetByName(SHEET_KEYS) || ss.insertSheet(SHEET_KEYS);
  if (k.getLastRow() === 0) {
    k.getRange(1, 1, 5, 2).setValues([
      ['Направление', 'Слова через запятую (можно начало слова: «активац» = активация, активации…)'],
      ['ТЛМ', 'дозвон, активац, онбординг, тлм, оборот'],
      ['Исходящие', 'холодн, база, базу, хантинг, партнёр, партнер, лидоген'],
      ['Входящие', 'заявк, входящ, банки.ру, квиз, marquiz'],
      ['Люди', 'Василий, Арина, Ксения, Степан']
    ]);
    k.setFrozenRows(1);
    k.getRange('1:1').setFontWeight('bold');
    k.setColumnWidth(1, 140);
    k.setColumnWidth(2, 520);
  }
  const def = ss.getSheetByName('Лист1') || ss.getSheetByName('Sheet1');
  if (def && ss.getSheets().length > 1 && def.getLastRow() === 0) ss.deleteSheet(def);

  const p = props();
  const res = [];
  res.push(tg('setWebhook', { url: p.getProperty('WEBAPP_URL'), drop_pending_updates: true, allowed_updates: ['message', 'callback_query'] }));
  if (p.getProperty('MINIAPP_URL')) {
    res.push(tg('setChatMenuButton', { menu_button: { type: 'web_app', text: 'Задачи', web_app: { url: p.getProperty('MINIAPP_URL') } } }));
  }
  res.push(tg('setMyCommands', { commands: [
    { command: 'today', description: 'План на сегодня' },
    { command: 'help', description: 'Как писать задачи' }
  ] }));

  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('morning').timeBased().everyDays(1).atHour(9).nearMinute(0).inTimezone(TZ).create();

  Logger.log(JSON.stringify(res));
}

/** Переподключить вебхук без потери ожидающих сообщений (если бот «залип»). */
function reconnect() {
  Logger.log(JSON.stringify(tg('setWebhook', { url: props().getProperty('WEBAPP_URL'), allowed_updates: ['message', 'callback_query'] })));
  const w = tg('getWebhookInfo', {}).result || {};
  Logger.log('ожидают доставки: ' + w.pending_update_count + (w.last_error_message ? ', последняя ошибка: ' + w.last_error_message : ''));
}

/* ───────────── Входная точка ───────────── */

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json({ ok: false, error: 'bad_json' }); }

  if (body.update_id !== undefined) {
    // Apps Script отвечает Телеграму редиректом, и тот иногда шлёт апдейт повторно — отсекаем дубли.
    const cache = CacheService.getScriptCache();
    const key = 'u' + body.update_id;
    if (!cache.get(key)) {
      cache.put(key, '1', 21600);
      try { handleUpdate(body); } catch (err) { console.error(err); }
    }
    // HtmlService отвечает Телеграму сразу кодом 200 (ContentService дал бы редирект 302, и Телеграм повторял бы апдейт).
    return HtmlService.createHtmlOutput('ok');
  }

  try { return json(handleApi(body)); } catch (err) { console.error(err); return json({ ok: false, error: String(err) }); }
}

function doGet() {
  return ContentService.createTextOutput('ok');
}

/* ───────────── Бот ───────────── */

function handleUpdate(u) {
  const p = props();
  const owner = p.getProperty('OWNER_ID');

  if (u.callback_query) {
    const cq = u.callback_query;
    if (String(cq.from.id) !== owner) return;
    handleCallback(cq);
    return;
  }

  const msg = u.message;
  if (!msg || !msg.from || msg.chat.type !== 'private') return;
  const from = String(msg.from.id);
  const text = (msg.text || msg.caption || '').trim();

  if (!owner) {
    if (/^\/start/.test(text)) {
      p.setProperty('OWNER_ID', from);
      send(from, helpText(true));
    }
    return;
  }
  if (from !== owner) return;

  if (!text) {
    send(from, 'Пока я понимаю только текст. Напиши задачу словами — или добавь подпись к картинке.');
    return;
  }
  if (/^\/start/.test(text) || /^\/help/.test(text)) { send(from, helpText(false)); return; }
  if (/^\/today/.test(text)) { const b = brief(false); send(from, b.text, b.kb); return; }

  let source = 'бот';
  if (msg.forward_origin) {
    const o = msg.forward_origin;
    const who = o.sender_user ? [o.sender_user.first_name, o.sender_user.last_name].filter(Boolean).join(' ')
      : o.sender_user_name || (o.chat && o.chat.title) || '';
    source = 'переслано' + (who ? ': ' + who : '');
  }

  withLock(() => {
    const t = createTask(text, source);
    send(from, confirmText(t), taskKb(t));
  });
}

function handleCallback(cq) {
  const [action, arg, id] = String(cq.data).split(':');
  const chat = cq.message.chat.id;
  const mid = cq.message.message_id;
  tg('answerCallbackQuery', { callback_query_id: cq.id });

  if (action === 'moveall') {
    const n = withLock(() => moveOverdueToToday());
    tg('editMessageReplyMarkup', { chat_id: chat, message_id: mid, reply_markup: openAppKb() });
    send(chat, n ? 'Перенесла на сегодня: ' + n + ' ' + plural(n, 'задачу', 'задачи', 'задач') + '.' : 'Просроченных нет.');
    return;
  }

  withLock(() => {
    let t = findTask(id);
    if (!t) { tg('editMessageText', { chat_id: chat, message_id: mid, text: 'Задача не найдена — возможно, её уже удалили.' }); return; }

    if (action === 'del') {
      deleteTask(id);
      tg('editMessageText', { chat_id: chat, message_id: mid, text: '✕ Удалено: ' + t.title });
      return;
    }
    if (action === 'due') t = updateTask(id, { due: arg === 'none' ? '' : resolveQuick(arg) });
    if (action === 'dir') t = updateTask(id, { dir: DIR_CODES[arg] || '' });
    if (action === 'idea') t = updateTask(id, { status: t.status === 'idea' ? 'inbox' : 'idea' });

    tg('editMessageText', { chat_id: chat, message_id: mid, text: confirmText(t), parse_mode: 'HTML', reply_markup: taskKb(t) });
  });
}

function helpText(first) {
  return (first ? 'Привет! Это твой рабочий таск-менеджер. Любое сообщение сюда — задача.\n\n' : '') +
    'Как писать:\n' +
    '• «отчёт по дозвону пт» — срок пятница, направление определю по словам\n' +
    '• сроки: сегодня, завтра, послезавтра, пн…вс, 15.11, через 3 дня, через неделю\n' +
    '• #тлм #исх #вх — указать направление явно\n' +
    '• @Арина — поручила, жду ответа\n' +
    '• «идея: …» — в копилку идей\n' +
    '• без срока и направления — во Входящие, разберёшь потом\n\n' +
    'Пересланные сообщения тоже становятся задачами. Список — кнопка «Задачи» слева от поля ввода.';
}

function confirmText(t) {
  return '✓ ' + esc(t.title) + '\n<i>' + esc(taskLabel(t)) + '</i>';
}

function taskLabel(t) {
  const parts = [];
  if (t.status === 'idea') parts.push('идея');
  else if (t.status === 'inbox') parts.push('входящие');
  else if (t.status === 'done') parts.push('готово');
  else parts.push(dueLabel(t.due));
  if (t.dir) parts.push(t.dir);
  if (t.person) parts.push('жду: ' + t.person);
  return parts.join(' · ');
}

function taskKb(t) {
  return { inline_keyboard: [
    [
      { text: 'Сегодня', callback_data: 'due:0:' + t.id },
      { text: 'Завтра', callback_data: 'due:1:' + t.id },
      { text: 'Пт', callback_data: 'due:fri:' + t.id },
      { text: 'Пн', callback_data: 'due:mon:' + t.id }
    ],
    [
      { text: (t.dir === 'ТЛМ' ? '● ' : '') + 'ТЛМ', callback_data: 'dir:t:' + t.id },
      { text: (t.dir === 'Исходящие' ? '● ' : '') + 'Исх', callback_data: 'dir:o:' + t.id },
      { text: (t.dir === 'Входящие' ? '● ' : '') + 'Вх', callback_data: 'dir:i:' + t.id }
    ],
    [
      { text: t.status === 'idea' ? 'Из идей в задачи' : 'В идеи', callback_data: 'idea:x:' + t.id },
      { text: 'Удалить', callback_data: 'del:x:' + t.id }
    ]
  ] };
}

function openAppKb() {
  const url = props().getProperty('MINIAPP_URL');
  return url ? { inline_keyboard: [[{ text: 'Открыть задачи', web_app: { url: url } }]] } : { inline_keyboard: [] };
}

/* ───────────── Утреннее сообщение ───────────── */

function morning() {
  const owner = props().getProperty('OWNER_ID');
  if (!owner) return;
  const dow = Number(Utilities.formatDate(new Date(), TZ, 'u')); // 1=пн … 7=вс
  if (dow > 5) return;
  const b = brief(true);
  if (b.empty) return; // нечего сказать — не беспокоим
  send(owner, b.text, b.kb);
}

function brief(isMorning) {
  const today = todayStr();
  const tasks = allTasks();
  const open = tasks.filter(t => t.status === 'active' || t.status === 'waiting');
  const overdue = open.filter(t => t.due && t.due < today);
  const todays = open.filter(t => t.due === today);
  const waitingOld = tasks.filter(t => t.status === 'waiting' && (!t.due || t.due >= today) && daysSince(t.created) >= 2);
  const inbox = tasks.filter(t => t.status === 'inbox');

  const L = [];
  L.push(isMorning ? '<b>Доброе утро!</b> ' + humanDate(today) : '<b>План на ' + humanDate(today) + '</b>');
  if (todays.length) {
    L.push('', '<b>Сегодня · ' + todays.length + '</b>');
    todays.slice(0, 8).forEach(t => L.push('• ' + esc(t.title) + (t.dir ? ' <i>' + esc(t.dir) + '</i>' : '')));
    if (todays.length > 8) L.push('… и ещё ' + (todays.length - 8));
  }
  if (overdue.length) {
    L.push('', '<b>Просрочено · ' + overdue.length + '</b>');
    overdue.slice(0, 5).forEach(t => L.push('• ' + esc(t.title) + ' <i>с ' + shortDate(t.due) + '</i>'));
    if (overdue.length > 5) L.push('… и ещё ' + (overdue.length - 5));
  }
  if (waitingOld.length) {
    L.push('', '<b>Жду ответа</b>');
    waitingOld.slice(0, 5).forEach(t => L.push('• ' + esc(t.person ? t.person + ' — ' : '') + esc(t.title) + ' <i>' + daysSince(t.created) + ' дн.</i>'));
  }
  if (inbox.length) L.push('', 'Во входящих: ' + inbox.length + ' — разбери, когда будет минутка.');

  const empty = !todays.length && !overdue.length && !waitingOld.length && !inbox.length;
  if (empty) L.push('', 'На сегодня задач нет.');

  const kb = openAppKb();
  if (overdue.length) kb.inline_keyboard.unshift([{ text: 'Перенести просроченные на сегодня', callback_data: 'moveall:x:x' }]);
  return { text: L.join('\n'), kb: kb, empty: empty };
}

function moveOverdueToToday() {
  const today = todayStr();
  let n = 0;
  allTasks().forEach(t => {
    if ((t.status === 'active' || t.status === 'waiting') && t.due && t.due < today) {
      sheet().getRange(t.row, COL.due).setValue(today);
      n++;
    }
  });
  return n;
}

/* ───────────── API для Mini App ───────────── */

function handleApi(body) {
  const user = verifyInitData(body.initData);
  if (!user || String(user.id) !== props().getProperty('OWNER_ID')) return { ok: false, error: 'auth' };

  const a = body.action;
  if (a === 'add') {
    const t = withLock(() => createTask(String(body.text || '').slice(0, 1000), 'приложение'));
    return listResponse({ added: t, label: taskLabel(t) });
  }
  if (a === 'update') {
    withLock(() => updateTask(body.id, body.fields || {}));
    return listResponse();
  }
  if (a === 'delete') {
    withLock(() => deleteTask(body.id));
    return listResponse();
  }
  if (a === 'moveall') {
    withLock(() => moveOverdueToToday());
    return listResponse();
  }
  return listResponse();
}

function listResponse(extra) {
  const today = todayStr();
  const tasks = allTasks()
    .filter(t => t.status !== 'done' || (t.doneAt || '').slice(0, 10) === today)
    .map(t => ({ id: t.id, created: t.created, title: t.title, dir: t.dir, person: t.person, due: t.due, status: t.status, doneAt: t.doneAt }));
  return Object.assign({ ok: true, today: today, dirs: directionNames(), tasks: tasks }, extra || {});
}

function verifyInitData(initData) {
  if (!initData) return null;
  const fields = {};
  let hash = '';
  String(initData).split('&').forEach(part => {
    const i = part.indexOf('=');
    if (i < 0) return;
    const k = part.slice(0, i);
    const v = decodeURIComponent(part.slice(i + 1).replace(/\+/g, '%20'));
    if (k === 'hash') hash = v; else fields[k] = v;
  });
  const dcs = Object.keys(fields).sort().map(k => k + '=' + fields[k]).join('\n');
  const secret = Utilities.computeHmacSha256Signature(Utilities.newBlob(token()).getBytes(), Utilities.newBlob('WebAppData').getBytes());
  const sig = Utilities.computeHmacSha256Signature(Utilities.newBlob(dcs).getBytes(), secret);
  const hex = sig.map(b => ('0' + (b & 255).toString(16)).slice(-2)).join('');
  if (hex !== hash) return null;
  if (Date.now() / 1000 - Number(fields.auth_date || 0) > 7 * 86400) return null;
  try { return JSON.parse(fields.user); } catch (e) { return null; }
}

/* ───────────── Разбор текста ───────────── */

function createTask(raw, source) {
  const parsed = parseTask(raw, todayStr(), loadKeys());
  const t = {
    id: Utilities.getUuid().replace(/-/g, '').slice(0, 8),
    created: nowStr(),
    title: parsed.title,
    dir: parsed.dir,
    person: parsed.person,
    due: parsed.due,
    status: parsed.status,
    doneAt: '',
    source: source || ''
  };
  sheet().appendRow([t.id, t.created, t.title, t.dir, t.person, t.due, t.status, t.doneAt, t.source]);
  return t;
}

/**
 * Чистая функция — без обращений к таблице, чтобы её можно было проверить отдельно.
 * keys = { dirs: [{name, words:[]}], people: [] }
 */
function parseTask(raw, today, keys) {
  let t = ' ' + String(raw).replace(/\s+/g, ' ').trim() + ' ';
  let status = '';
  let dir = '';
  let person = '';
  let due = '';

  const ideaRe = /^\s*(?:#?идея|💡)\s*[:\-—–.]?\s*/i;
  if (ideaRe.test(t)) { status = 'idea'; t = ' ' + t.replace(ideaRe, ''); }

  t = t.replace(/(^|\s)#([0-9a-zа-яё_.]+)/gi, (m, pre, tag) => {
    if (/^идея$/i.test(tag)) { status = 'idea'; return pre; }
    const d = tagToDir(tag, keys);
    if (d) { dir = d; return pre; }
    return m;
  });

  t = t.replace(/(^|\s)@([a-zа-яё][0-9a-zа-яё_]*)/gi, (m, pre, name) => {
    person = normalizePerson(name, keys);
    return pre;
  });

  const d = parseDue(t, today);
  if (d) { due = d.due; t = d.rest; }

  if (!dir) dir = guessDir(t, keys);

  let title = t.replace(/\s+/g, ' ').trim()
    .replace(/^[\s,;:\-—–.]+/, '').replace(/[\s,;:\-—–]+$/, '').trim();
  if (!title) title = String(raw).trim();
  title = title.charAt(0).toUpperCase() + title.slice(1);

  if (!status) {
    if (person) status = 'waiting';
    else if (!due && !dir) status = 'inbox';
    else status = 'active';
  }
  if (status === 'idea') due = '';
  return { title: title, dir: dir, person: person, due: due, status: status };
}

const B = '(^|[\\s,(])';
const E = '(?=[\\s,.!?;:)]|$)';
const PRE = '(?:(?:до|к|ко|на|в|во|по)\\s+)?';

function parseDue(t, today) {
  const rules = [
    { re: 'послезавтра', f: () => addDays(today, 2) },
    { re: 'завтра', f: () => addDays(today, 1) },
    { re: 'сегодня|сёдня', f: () => today },
    { re: 'через\\s+неделю', f: () => addDays(today, 7) },
    { re: 'через\\s+(\\d{1,2})\\s+недел[июь]', f: m => addDays(today, 7 * Number(m[2])) },
    { re: 'через\\s+(\\d{1,3})\\s+(?:день|дня|дней|дн\\.?)', f: m => addDays(today, Number(m[2])) },
    { re: 'через\\s+день', f: () => addDays(today, 2) },
    { re: 'конц[ауе]\\s+недели|на\\s+этой\\s+неделе|на\\s+неделе', f: () => nextDow(today, 5) },
    { re: 'понедельник[а-я]*|пн', f: () => nextDow(today, 1) },
    { re: 'вторник[а-я]*|вт', f: () => nextDow(today, 2) },
    { re: 'сред[уаые]|ср', f: () => nextDow(today, 3) },
    { re: 'четверг[а-я]*|чт', f: () => nextDow(today, 4) },
    { re: 'пятниц[уаые]|пт', f: () => nextDow(today, 5) },
    { re: 'суббот[уаые]|сб', f: () => nextDow(today, 6) },
    { re: 'воскресень[еяю]|вс', f: () => nextDow(today, 0) },
    { re: '(\\d{1,2})[./](\\d{1,2})(?:[./](\\d{2,4}))?', f: m => explicitDate(m, today) }
  ];
  for (let i = 0; i < rules.length; i++) {
    const rx = new RegExp(B + PRE + '(?:' + rules[i].re + ')' + E, 'i');
    const m = t.match(rx);
    if (!m) continue;
    // для даты числами группы сдвинуты: m[1] — граница, дальше день/месяц/год
    const due = rules[i].f(m);
    if (!due) continue;
    return { due: due, rest: t.replace(rx, (all, pre) => pre + ' ') };
  }
  return null;
}

function explicitDate(m, today) {
  const day = Number(m[2]);
  const mon = Number(m[3]);
  if (!day || !mon || day > 31 || mon > 12) return '';
  let year = m[4] ? Number(m[4]) : Number(today.slice(0, 4));
  if (year < 100) year += 2000;
  let s = year + '-' + pad(mon) + '-' + pad(day);
  if (!m[4] && s < addDays(today, -60)) s = (year + 1) + '-' + pad(mon) + '-' + pad(day);
  return s;
}

function tagToDir(tag, keys) {
  const t = tag.toLowerCase();
  if (/^(тлм|tlm|телемаркетинг)$/.test(t)) return 'ТЛМ';
  if (/^(исх|исходящие|исходящ)$/.test(t)) return 'Исходящие';
  if (/^(вх|входящие|входящ)$/.test(t)) return 'Входящие';
  const custom = (keys.dirs || []).filter(d => d.name.toLowerCase().indexOf(t) === 0)[0];
  return custom ? custom.name : '';
}

function guessDir(text, keys) {
  const low = ' ' + text.toLowerCase().replace(/ё/g, 'е') + ' ';
  let best = '', bestN = 0;
  (keys.dirs || []).forEach(d => {
    let n = 0;
    d.words.forEach(w => {
      const ww = w.toLowerCase().replace(/ё/g, 'е');
      if (!ww) return;
      const rx = new RegExp('(^|[^0-9a-zа-я])' + ww.replace(/[.*+?^$(){}|[\]\\]/g, '\\$&'), 'i');
      if (rx.test(low)) n++;
    });
    if (n > bestN) { best = d.name; bestN = n; }
  });
  return best;
}

function normalizePerson(name, keys) {
  const low = name.toLowerCase();
  const hit = (keys.people || []).filter(p => p.toLowerCase().indexOf(low) === 0 || low.indexOf(p.toLowerCase()) === 0)[0];
  return hit || name.charAt(0).toUpperCase() + name.slice(1);
}

function loadKeys() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_KEYS);
  const out = { dirs: [], people: [] };
  if (!sh || sh.getLastRow() < 2) return out;
  sh.getRange(2, 1, sh.getLastRow() - 1, 2).getDisplayValues().forEach(r => {
    const name = String(r[0]).trim();
    const words = String(r[1]).split(',').map(s => s.trim()).filter(Boolean);
    if (!name) return;
    if (/^люди$/i.test(name)) out.people = words;
    else out.dirs.push({ name: name, words: words });
  });
  return out;
}

function directionNames() {
  const names = loadKeys().dirs.map(d => d.name);
  return names.length ? names : ['ТЛМ', 'Исходящие', 'Входящие'];
}

/* ───────────── Таблица ───────────── */

function sheet() { return SpreadsheetApp.getActive().getSheetByName(SHEET_TASKS); }

function allTasks() {
  const sh = sheet();
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, HEAD.length).getDisplayValues()
    .map((r, i) => ({ row: i + 2, id: r[0], created: r[1], title: r[2], dir: r[3], person: r[4], due: r[5], status: r[6], doneAt: r[7], source: r[8] }))
    .filter(t => t.id);
}

function findTask(id) { return allTasks().filter(t => t.id === String(id))[0] || null; }

function updateTask(id, f) {
  const t = findTask(id);
  if (!t) return null;
  const n = Object.assign({}, t);

  if ('title' in f && String(f.title).trim()) n.title = String(f.title).trim().slice(0, 500);
  if ('dir' in f) n.dir = String(f.dir || '');
  if ('person' in f) n.person = String(f.person || '');
  if ('due' in f) n.due = /^\d{4}-\d{2}-\d{2}$/.test(f.due) ? f.due : '';
  if ('status' in f && STATUS_RU[f.status]) n.status = f.status;

  // переходы статусов
  if (!('status' in f)) {
    if ((n.status === 'inbox' || n.status === 'idea') && ('due' in f && n.due)) n.status = n.person ? 'waiting' : 'active';
    else if (n.status === 'inbox' && 'dir' in f && n.dir) n.status = 'active';
  }
  if (n.status === 'idea') n.due = '';
  if (n.status === 'done' && t.status !== 'done') n.doneAt = nowStr();
  if (n.status !== 'done') n.doneAt = '';

  sheet().getRange(t.row, 1, 1, HEAD.length).setValues([[n.id, n.created, n.title, n.dir, n.person, n.due, n.status, n.doneAt, n.source]]);
  return n;
}

function deleteTask(id) {
  const t = findTask(id);
  if (t) sheet().deleteRow(t.row);
}

/* ───────────── Даты и мелочи ───────────── */

function todayStr() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function nowStr() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'); }
function pad(n) { return ('0' + n).slice(-2); }
function toD(s) { return new Date(Date.UTC(+s.slice(0, 4), +s.slice(5, 7) - 1, +s.slice(8, 10), 12)); }
function fromD(d) { return d.toISOString().slice(0, 10); }
function addDays(s, n) { const d = toD(s); d.setUTCDate(d.getUTCDate() + n); return fromD(d); }
function nextDow(s, dow) { const cur = toD(s).getUTCDay(); return addDays(s, (dow - cur + 7) % 7); }
function diffDays(a, b) { return Math.round((toD(b) - toD(a)) / 86400000); }
function daysSince(created) { return created ? diffDays(created.slice(0, 10), todayStr()) : 0; }
function shortDate(s) { return s.slice(8, 10) + '.' + s.slice(5, 7); }
function humanDate(s) { const d = toD(s); return d.getUTCDate() + ' ' + MONTHS[d.getUTCMonth()]; }

function resolveQuick(arg) {
  const today = todayStr();
  if (arg === 'fri') return nextDow(today, 5);
  if (arg === 'mon') { const d = nextDow(today, 1); return d === today ? addDays(today, 7) : d; }
  return addDays(today, Number(arg) || 0);
}

function dueLabel(due) {
  if (!due) return 'без срока';
  const today = todayStr();
  const n = diffDays(today, due);
  if (n < 0) return 'просрочено, ' + shortDate(due);
  if (n === 0) return 'сегодня';
  if (n === 1) return 'завтра';
  if (n < 7) return WD_SHORT[toD(due).getUTCDay()] + ', ' + shortDate(due);
  return shortDate(due);
}

function plural(n, one, few, many) {
  const a = n % 10, b = n % 100;
  if (a === 1 && b !== 11) return one;
  if (a >= 2 && a <= 4 && (b < 12 || b > 14)) return few;
  return many;
}

function esc(s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }
function props() { return PropertiesService.getScriptProperties(); }
function token() { return props().getProperty('BOT_TOKEN'); }
function json(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function withLock(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function send(chat, text, kb) {
  const p = { chat_id: chat, text: text, parse_mode: 'HTML', disable_web_page_preview: true };
  if (kb) p.reply_markup = kb;
  return tg('sendMessage', p);
}

function tg(method, payload) {
  const r = UrlFetchApp.fetch('https://api.telegram.org/bot' + token() + '/' + method, {
    method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true
  });
  return JSON.parse(r.getContentText());
}
