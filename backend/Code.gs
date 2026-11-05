/**
 * Mirror · онлайн-витрина · сервер заказов (Google Apps Script)
 *
 * Что делает:
 *  - отдаёт сайту каталог, зоны доставки и занятые интервалы из Google-таблицы;
 *  - принимает заказ с сайта, пишет его в «Заказы» и обновляет «Клиенты»;
 *  - присылает заказ владельцу в Telegram с кнопками статусов;
 *  - сообщает покупателю о смене статуса, если он подключил бота.
 *
 * Установка: см. backend/README.md
 */

// ===== НАСТРОЙКИ (заполняются один раз) =====
const BOT_TOKEN = 'ВСТАВЬТЕ_ТОКЕН_БОТА';
const ADMIN_KEY = 'ВСТАВЬТЕ_КОД_ВЛАДЕЛЬЦА';   // секрет в ссылке для подключения владельца
const HOOK_KEY  = 'ВСТАВЬТЕ_КЛЮЧ_ВЕБХУКА';    // секрет в адресе вебхука Telegram
const SITE_URL  = 'https://seni4jap.github.io/flowers-demo/';
const TZ = 'Europe/Moscow';

// ===== СТРУКТУРА ТАБЛИЦЫ =====
const SH = { cat: 'Каталог', ord: 'Заказы', cli: 'Клиенты', set: 'Настройки' };
const CAT_H = ['ID', 'Название', 'Категория', 'Цена (размер M), ₽', 'Состав', 'Фото', 'Фото 2', 'В наличии'];
const ORD_H = ['№', 'Создан', 'Статус', 'Букет', 'Размер', 'Букет, ₽', 'Доставка, ₽', 'Итого, ₽',
  'Дата', 'Интервал', 'Получение', 'Адрес', 'Покупатель', 'Телефон', 'Получатель', 'Телефон получателя',
  'Открытка', 'Фото до отправки', 'Анонимно', 'Комментарий', 'Взял в работу', 'Обновлён',
  'Telegram покупателя', 'тех: сообщения', 'тех: ключ'];
const CLI_H = ['Телефон', 'Имя', 'Заказов', 'Сумма, ₽', 'Первый заказ', 'Последний заказ', 'Последний букет', 'Telegram'];
const O = {}; ORD_H.forEach(function (h, i) { O[h] = i; });

const CATS = { 'Авторские': 'author', 'В коробке': 'box', 'Корзины': 'basket', 'Моно': 'mono' };
const SIZES = { S: 0.75, M: 1, L: 1.4 };
const SLOTS = [[9, 12], [12, 15], [15, 18], [18, 21]];
const ST = {
  new: 'Новый', work: 'В работе', ready: 'Собран', way: 'В пути',
  done: 'Доставлен', given: 'Выдан', cancel: 'Отменён'
};
const ST_ICON = { new: '🆕', work: '🛠', ready: '💐', way: '🚚', done: '✅', given: '✅', cancel: '✖️' };
const MONTHS = ['января', 'февраля', 'марта', 'апреля', 'мая', 'июня', 'июля', 'августа', 'сентября', 'октября', 'ноября', 'декабря'];
const WD = ['вс', 'пн', 'вт', 'ср', 'чт', 'пт', 'сб'];

const SEED_CATALOG = [
  [1, 'Нежность пионов', 'Авторские', 8500, 'Пионовидные розы, гортензия, эустома', 'img/hero.webp', 'img/c8.webp', 'Да'],
  [2, 'Розовый сад', 'Авторские', 6900, 'Розы, эустома, маттиола', 'img/c2.webp', 'img/c1.webp', 'Да'],
  [3, '101 белая роза', 'Корзины', 18900, 'Корзина, белые розы, эвкалипт', 'img/n_white1.webp', 'img/n_white2.webp', 'Да'],
  [4, 'Пудровое облако', 'В коробке', 9800, 'Розы, гипсофила, эвкалипт', 'img/c3.webp', 'img/c6.webp', 'Да'],
  [5, 'Белые гортензии', 'Моно', 7900, 'Гортензия, пионы, зелень', 'img/c7.webp', 'img/c2.webp', 'Да'],
  [6, 'Кустовые розы в корзине', 'Корзины', 21500, 'Пионовидные кустовые розы', 'img/n_pink1.webp', 'img/n_pink2.webp', 'Да'],
  [7, 'Лилии в шляпной коробке', 'В коробке', 11900, 'Лилии, розы, альстромерия', 'img/c4.webp', 'img/c3.webp', 'Да'],
  [8, 'Сиреневый вечер', 'В коробке', 14500, 'Гортензия, розы, орхидея', 'img/c6.webp', 'img/c4.webp', 'Да']
];
const SEED_SETTINGS = [
  ['Магазин', 'Pivoine'],
  ['Телефон магазина', '+7 999 000-00-00'],
  ['Самовывоз', 'м. Белорусская'],
  ['Лимит заказов на интервал', 3],
  ['Зона: В пределах МКАД', 0],
  ['Зона: До 10 км от МКАД', 700],
  ['Зона: До 30 км от МКАД', 1500],
  ['Адрес веб-приложения', '']
];

// ===== УСТАНОВКА =====

/** Шаг 1. Создаёт листы таблицы и проверяет бота. Запускать один раз. */
function setup() {
  const ss = SpreadsheetApp.getActive();
  makeSheet_(ss, SH.cat, CAT_H, SEED_CATALOG, [50, 220, 120, 150, 300, 160, 160, 100]);
  makeSheet_(ss, SH.ord, ORD_H, [], null);
  makeSheet_(ss, SH.cli, CLI_H, [], [170, 160, 90, 110, 140, 140, 200, 140]);
  makeSheet_(ss, SH.set, ['Параметр', 'Значение'], SEED_SETTINGS, [260, 420]);
  const ord = ss.getSheetByName(SH.ord);
  ord.hideColumns(O['тех: сообщения'] + 1, 2);
  ss.setSpreadsheetTimeZone(TZ);
  styleAll_();
  buildSummary_();

  const me = tg_('getMe', {});
  if (!me || !me.ok) throw new Error('Бот не отвечает. Проверьте BOT_TOKEN в начале файла.');
  props_().setProperty('BOT', me.result.username);
  Logger.log('Готово. Бот: @' + me.result.username);
  Logger.log('Ссылка владельца: https://t.me/' + me.result.username + '?start=admin_' + ADMIN_KEY);
}

/** Шаг 2. После «Развернуть → Веб-приложение»: вставьте адрес в «Настройки» и запустите. */
function connectBot() {
  const url = String(settings_()['Адрес веб-приложения'] || '').trim();
  if (!/^https:\/\/script\.google\.com\/.+\/exec$/.test(url)) {
    throw new Error('В листе «Настройки» в строке «Адрес веб-приложения» должен быть адрес, который заканчивается на /exec');
  }
  const r = tg_('setWebhook', {
    url: url + '?k=' + HOOK_KEY,
    allowed_updates: ['message', 'callback_query'],
    drop_pending_updates: true
  });
  Logger.log(JSON.stringify(r));
  if (!r || !r.ok) throw new Error('Telegram не принял вебхук: ' + JSON.stringify(r));
  Logger.log('Бот подключён. Откройте ссылку владельца: https://t.me/' + botName_() + '?start=admin_' + ADMIN_KEY);
}

function makeSheet_(ss, name, head, rows, widths) {
  let sh = ss.getSheetByName(name);
  if (sh && sh.getLastRow() > 0) return sh;          // не трогаем уже заполненный лист
  if (!sh) sh = ss.insertSheet(name);
  sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#EFE3BC');
  sh.setFrozenRows(1);
  if (rows.length) sh.getRange(2, 1, rows.length, head.length).setValues(rows);
  if (widths) widths.forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  return sh;
}

// ===== ВХОДНЫЕ ТОЧКИ =====

function doGet(e) {
  const a = (e && e.parameter && e.parameter.a) || '';
  try {
    if (a === 'init') return json_(init_());
    return json_({ ok: true, service: 'mirror-orders' });
  } catch (err) {
    return json_({ ok: false, err: 'server', msg: String(err) });
  }
}

function doPost(e) {
  let body = {};
  try { body = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (err) { body = {}; }

  // Обновление от Telegram. Отвечаем HtmlService: он отдаёт 200 без редиректа,
  // иначе Telegram считает ответ ошибкой и присылает обновление повторно.
  if (body.update_id !== undefined) {
    if (e.parameter && e.parameter.k === HOOK_KEY) {
      try { handleUpdate_(body); } catch (err) { console.error(err); }
    }
    return HtmlService.createHtmlOutput('ok');
  }

  try {
    if (body.a === 'order') return json_(createOrder_(body));
    return json_({ ok: false, err: 'unknown' });
  } catch (err) {
    console.error(err);
    return json_({ ok: false, err: 'server' });
  }
}

// ===== ДАННЫЕ ДЛЯ САЙТА =====

function init_() {
  const st = settings_();
  return {
    ok: true,
    catalog: catalog_().filter(function (x) { return x.stock; }).map(function (x) {
      return { id: x.id, n: x.n, c: x.c, p: x.p, s: x.s, img: x.img, img2: x.img2 };
    }),
    zones: zones_(st),
    full: fullSlots_(st),
    pickup: st['Самовывоз'] || '',
    phone: st['Телефон магазина'] || '',
    bot: botName_()
  };
}

function catalog_() {
  const v = sheet_(SH.cat).getDataRange().getValues().slice(1);
  return v.filter(function (r) { return r[0] !== '' && r[1] !== ''; }).map(function (r) {
    return {
      id: String(r[0]), n: String(r[1]), c: CATS[String(r[2]).trim()] || 'author',
      p: Number(String(r[3]).replace(/[^\d.]/g, '')) || 0, s: String(r[4] || ''),
      img: String(r[5] || ''), img2: String(r[6] || r[5] || ''),
      stock: !/^(нет|no|0|false)$/i.test(String(r[7]).trim())
    };
  });
}

function settings_() {
  const out = {};
  sheet_(SH.set).getDataRange().getValues().slice(1).forEach(function (r) {
    if (r[0] !== '') out[String(r[0]).trim()] = r[1];
  });
  return out;
}

function zones_(st) {
  return Object.keys(st).filter(function (k) { return k.indexOf('Зона:') === 0; }).map(function (k) {
    return { n: k.replace('Зона:', '').trim(), p: Number(st[k]) || 0 };
  });
}

/** Ключи «2026-11-05|12» для интервалов, где заказов уже не меньше лимита. */
function fullSlots_(st) {
  const limit = Number(st['Лимит заказов на интервал']) || 0;
  if (!limit) return [];
  const today = isoToday_(), cnt = {};
  ordersRows_().forEach(function (r) {
    if (String(r[O['Статус']]) === ST.cancel) return;
    const k = slotKey_(r);
    if (k && k.slice(0, 10) >= today) cnt[k] = (cnt[k] || 0) + 1;
  });
  return Object.keys(cnt).filter(function (k) { return cnt[k] >= limit; });
}

function slotKey_(r) {
  const d = d2s_(r[O['Дата']], 'dd.MM.yyyy').match(/(\d{2})\.(\d{2})\.(\d{4})/);
  const t = String(r[O['Интервал']]).match(/^(\d{1,2})/);
  if (!d || !t) return '';
  return d[3] + '-' + d[2] + '-' + d[1] + '|' + Number(t[1]);
}

// ===== ЗАКАЗ =====

function createOrder_(d) {
  const st = settings_();
  const item = catalog_().filter(function (x) { return x.id === String(d.id); })[0];
  if (!item || !item.stock) return { ok: false, err: 'stock' };
  if (!SIZES[d.size]) return { ok: false, err: 'bad', field: 'size' };
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(d.date)) || d.date < isoToday_()) return { ok: false, err: 'bad', field: 'date' };
  const slot = SLOTS.filter(function (s) { return s[0] === Number(d.slot); })[0];
  if (!slot) return { ok: false, err: 'bad', field: 'slot' };
  const pickup = d.mode === 'p';
  const zones = zones_(st), zone = zones[Number(d.zone)] || zones[0] || { n: '', p: 0 };
  const street = clean_(d.street, 200);
  if (!pickup && street.length < 5) return { ok: false, err: 'bad', field: 'street' };
  const name = clean_(d.name, 60);
  const phone = phoneFmt_(d.phone);
  if (name.length < 2) return { ok: false, err: 'bad', field: 'name' };
  if (!phone) return { ok: false, err: 'bad', field: 'phone' };
  const other = d.who === 'other';

  const bouquet = r100_(item.p * SIZES[d.size]);
  const delivery = pickup ? 0 : zone.p;
  const total = bouquet + delivery;
  const dateVal = new Date(Number(d.date.slice(0, 4)), Number(d.date.slice(5, 7)) - 1, Number(d.date.slice(8, 10)), 12);
  const interval = slot[0] + ':00–' + slot[1] + ':00';

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  let no, row, key;
  try {
    const limit = Number(st['Лимит заказов на интервал']) || 0;
    const rows = ordersRows_();
    if (limit) {
      const want = d.date + '|' + slot[0];
      const busy = rows.filter(function (r) {
        return String(r[O['Статус']]) !== ST.cancel && slotKey_(r) === want;
      }).length;
      if (busy >= limit) return { ok: false, err: 'slot_full' };
    }
    no = rows.reduce(function (m, r) { return Math.max(m, Number(r[O['№']]) || 0); }, 1000) + 1;
    key = Utilities.getUuid().replace(/-/g, '').slice(0, 10);
    const rec = [];
    rec[O['№']] = no;
    rec[O['Создан']] = new Date();
    rec[O['Статус']] = ST.new;
    rec[O['Букет']] = txt_(item.n);
    rec[O['Размер']] = d.size;
    rec[O['Букет, ₽']] = bouquet;
    rec[O['Доставка, ₽']] = delivery;
    rec[O['Итого, ₽']] = total;
    rec[O['Дата']] = dateVal;
    rec[O['Интервал']] = txt_(interval);
    rec[O['Получение']] = pickup ? 'Самовывоз' : 'Курьер';
    rec[O['Адрес']] = txt_(pickup ? 'Самовывоз: ' + (st['Самовывоз'] || '') : street + ' (' + zone.n + ')');
    rec[O['Покупатель']] = txt_(name);
    rec[O['Телефон']] = txt_(phone);
    rec[O['Получатель']] = txt_(other ? clean_(d.rname, 60) : 'Сам покупатель');
    rec[O['Телефон получателя']] = txt_(other ? (phoneFmt_(d.rphone) || clean_(d.rphone, 30)) : '');
    rec[O['Открытка']] = txt_(d.card ? clean_(d.cardText, 200) || 'Да, без текста' : '');
    rec[O['Фото до отправки']] = d.photo ? 'Да' : 'Нет';
    rec[O['Анонимно']] = d.anon ? 'Да' : 'Нет';
    rec[O['Комментарий']] = txt_(clean_(d.comment, 500));
    rec[O['Взял в работу']] = '';
    rec[O['Обновлён']] = new Date();
    rec[O['Telegram покупателя']] = '';
    rec[O['тех: сообщения']] = '[]';
    rec[O['тех: ключ']] = key;
    const sh = sheet_(SH.ord);
    sh.appendRow(rec);
    row = sh.getLastRow();
    upsertClient_(phone, name, total, item.n + ', ' + d.size, '');
  } finally {
    lock.releaseLock();
  }

  // Уведомление владельцам
  const o = orderAt_(row);
  const sent = [];
  owners_().forEach(function (chat) {
    const r = tg_('sendMessage', {
      chat_id: chat, text: ownerText_(o), parse_mode: 'HTML',
      reply_markup: { inline_keyboard: buttons_(o) }, disable_web_page_preview: true
    });
    if (r && r.ok) sent.push([chat, r.result.message_id]);
  });
  sheet_(SH.ord).getRange(row, O['тех: сообщения'] + 1).setValue(JSON.stringify(sent));

  // Заказ из Telegram (сайт открыт кнопкой «Каталог» в боте): привязываем покупателя сразу
  let linked = false;
  const tgUser = verifyWebApp_(d.tg);
  if (tgUser) {
    linked = linkBuyer_(tgUser.id, tgUser, no, key);
  }

  return { ok: true, no: no, total: total, key: key, bot: botName_(), linked: linked };
}

function upsertClient_(phone, name, sum, last, tgId) {
  const sh = sheet_(SH.cli);
  const v = sh.getDataRange().getValues();
  const digits = phone.replace(/\D/g, '');
  const today = new Date();
  for (let i = 1; i < v.length; i++) {
    if (String(v[i][0]).replace(/\D/g, '') === digits) {
      if (sum) {
        sh.getRange(i + 1, 2, 1, 6).setValues([[txt_(name), (Number(v[i][2]) || 0) + 1,
          (Number(v[i][3]) || 0) + sum, v[i][4] || today, today, txt_(last)]]);
      }
      if (tgId) sh.getRange(i + 1, 8).setValue(txt_(tgId));
      return;
    }
  }
  if (sum) sh.appendRow([txt_(phone), txt_(name), 1, sum, today, today, txt_(last), txt_(tgId || '')]);
}

// ===== TELEGRAM =====

function handleUpdate_(u) {
  const cache = CacheService.getScriptCache();
  if (cache.get('u' + u.update_id)) return;
  cache.put('u' + u.update_id, '1', 21600);

  if (u.callback_query) return onButton_(u.callback_query);
  const m = u.message;
  if (!m || !m.chat) return;
  const chat = m.chat.id, text = String(m.text || '').trim();
  const isOwner = owners_().indexOf(String(chat)) >= 0;

  if (isOwner && m.photo && m.photo.length) return forwardPhoto_(chat, m);

  if (text === '/start admin_' + ADMIN_KEY) {
    addOwner_(chat);
    return say_(chat, '✅ Готово! Теперь заказы с сайта будут приходить сюда.\n\n' +
      'Под каждым заказом — кнопки статусов. Покупатель, подключивший бота, получит уведомление о каждом шаге.\n\n' +
      '/today — заказы на сегодня\n/tomorrow — на завтра');
  }

  const mo = text.match(/^\/start o_(\d+)_([A-Za-z0-9]+)$/);
  if (mo) return linkBuyer_(chat, m.from, Number(mo[1]), mo[2]);

  if (isOwner && (text === '/today' || text === '/tomorrow')) return dayList_(chat, text === '/today' ? 0 : 1);

  const shop = settings_()['Магазин'] || 'цветочной мастерской';
  if (text === '/catalog') {
    return tg_('sendMessage', { chat_id: chat, text: 'Каталог ' + esc_(shop) + ' — выбирайте букет 🌸', parse_mode: 'HTML', reply_markup: catalogButton_() });
  }
  tg_('sendMessage', {
    chat_id: chat, parse_mode: 'HTML', reply_markup: catalogButton_(), disable_web_page_preview: true,
    text: 'Здравствуйте! Это цветочная мастерская <b>' + esc_(shop) + '</b> 🌸\n\n' +
      'Соберём букет под ваш повод и привезём по Москве за 2 часа. Каталог с ценами — по кнопке ниже.\n\n' +
      'После заказа сюда будут приходить статусы: собран, передан курьеру, доставлен.'
  });
}

function linkBuyer_(chat, from, no, key) {
  const row = findOrder_(no);
  if (!row) return say_(chat, 'Не нашли такой заказ. Позвоните нам, пожалуйста.');
  const sh = sheet_(SH.ord);
  const r = sh.getRange(row, 1, 1, ORD_H.length).getValues()[0];
  if (String(r[O['тех: ключ']]) !== key) return say_(chat, 'Ссылка устарела. Позвоните нам, пожалуйста.');
  const handle = (from && from.username) ? '@' + from.username : String(chat);
  const o = orderAt_(row);
  const r2 = say_(chat, 'Здравствуйте, ' + esc_(o.name.split(' ')[0]) + '! 🌸\n\n' +
    'Ваш заказ <b>№' + o.no + '</b>\n' + esc_(o.item) + ', размер ' + o.size + '\n' +
    whenText_(o) + '\nИтого: ' + money_(o.total) + '\n\n' +
    'Статус: ' + ST_ICON[o.code] + ' ' + o.status + '\n\n' +
    (o.photo ? '📷 Перед отправкой пришлём сюда фото вашего букета.\n' : '') +
    'Сюда придёт сообщение, когда статус изменится.');
  if (!r2 || !r2.ok) return false;          // человек не запускал бота — писать ему нельзя
  sh.getRange(row, O['Telegram покупателя'] + 1).setValue(txt_(chat + ' ' + handle));
  upsertClient_(String(r[O['Телефон']]), String(r[O['Покупатель']]), 0, '', handle);
  refreshOwners_(row);                     // в карточке владельца появится «🔔 получает статусы»
  return true;
}

/** Проверка подписи данных Telegram Mini App (initData). Возвращает пользователя или null. */
function verifyWebApp_(initData) {
  if (!initData || typeof initData !== 'string' || initData.length > 4096) return null;
  try {
    let hash = '';
    const kv = [];
    initData.split('&').forEach(function (p) {
      const i = p.indexOf('=');
      if (i < 0) return;
      const k = p.slice(0, i), v = decodeURIComponent(p.slice(i + 1));
      if (k === 'hash') hash = v; else kv.push(k + '=' + v);
    });
    if (!hash) return null;
    kv.sort();
    const secret = Utilities.computeHmacSha256Signature(BOT_TOKEN, 'WebAppData');
    const sig = Utilities.computeHmacSha256Signature(Utilities.newBlob(kv.join('\n')).getBytes(), secret);
    const hex = sig.map(function (b) { return ('0' + ((b + 256) % 256).toString(16)).slice(-2); }).join('');
    if (hex !== hash) return null;
    const f = {};
    kv.forEach(function (x) { const i = x.indexOf('='); f[x.slice(0, i)] = x.slice(i + 1); });
    if (Date.now() / 1000 - Number(f.auth_date || 0) > 86400) return null;
    const u = JSON.parse(f.user || 'null');
    return u && u.id ? u : null;
  } catch (e) {
    console.error(e);
    return null;
  }
}

/** Флорист ответил фотографией на карточку заказа — пересылаем покупателю. */
function forwardPhoto_(chat, m) {
  const reply = m.reply_to_message;
  if (!reply) {
    return say_(chat, '📷 Чтобы отправить фото покупателю, <b>ответьте этой фотографией на карточку заказа</b>: ' +
      'смахните карточку влево (или «Ответить») и прикрепите фото.');
  }
  const v = sheet_(SH.ord).getDataRange().getValues();
  let row = 0;
  for (let i = 1; i < v.length && !row; i++) {
    let msgs = [];
    try { msgs = JSON.parse(String(v[i][O['тех: сообщения']] || '[]')); } catch (e) { msgs = []; }
    msgs.forEach(function (cm) { if (String(cm[0]) === String(chat) && Number(cm[1]) === reply.message_id) row = i + 1; });
  }
  if (!row) return say_(chat, 'Не нашёл заказ для этой карточки. Ответьте фото именно на сообщение с заказом.');
  const o = orderAt_(row), shop = settings_();
  if (!o.buyerChat) {
    return say_(chat, '📷 Покупатель заказа №' + o.no + ' не подключил Telegram. Отправьте фото в WhatsApp: ' + esc_(o.phone));
  }
  const ph = m.photo[m.photo.length - 1].file_id;
  const r = tg_('sendPhoto', {
    chat_id: o.buyerChat, photo: ph, parse_mode: 'HTML',
    caption: 'Ваш букет <b>№' + o.no + '</b> готов 🌸\n' + esc_(o.item) + ', размер ' + o.size + '\n\n' +
      'Если хотите что-то поправить — позвоните нам: ' + esc_(shop['Телефон магазина'] || '') + '.'
  });
  if (!r || !r.ok) return say_(chat, 'Не получилось отправить фото покупателю. Отправьте в WhatsApp: ' + esc_(o.phone));
  sheet_(SH.ord).getRange(row, O['Фото до отправки'] + 1).setValue('Отправлено');
  refreshOwners_(row);
  tg_('sendMessage', { chat_id: chat, text: '✅ Фото отправлено покупателю (заказ №' + o.no + ')', reply_to_message_id: m.message_id });
}

function onButton_(q) {
  const chat = q.message && q.message.chat && q.message.chat.id;
  if (owners_().indexOf(String(chat)) < 0) return tg_('answerCallbackQuery', { callback_query_id: q.id, text: 'Нет доступа' });
  const p = String(q.data || '').split(':');            // s:<№>:<код>
  if (p[0] !== 's' || !ST[p[2]]) return tg_('answerCallbackQuery', { callback_query_id: q.id });
  const no = Number(p[1]), code = p[2];

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  let row, o;
  try {
    row = findOrder_(no);
    if (!row) return tg_('answerCallbackQuery', { callback_query_id: q.id, text: 'Заказ не найден' });
    o = orderAt_(row);
    if (next_(o).indexOf(code) < 0) {
      tg_('answerCallbackQuery', { callback_query_id: q.id, text: 'Статус уже: ' + o.status });
      refreshOwners_(row);
      return;
    }
    const sh = sheet_(SH.ord);
    sh.getRange(row, O['Статус'] + 1).setValue(ST[code]);
    if (code === 'work') sh.getRange(row, O['Взял в работу'] + 1).setValue(txt_(personName_(q.from)));
    sh.getRange(row, O['Обновлён'] + 1).setValue(new Date());
  } finally {
    lock.releaseLock();
  }
  tg_('answerCallbackQuery', { callback_query_id: q.id, text: ST_ICON[code] + ' ' + ST[code] });
  refreshOwners_(row);
  notifyBuyer_(orderAt_(row));
}

function refreshOwners_(row) {
  const o = orderAt_(row);
  o.msgs.forEach(function (cm) {
    tg_('editMessageText', {
      chat_id: cm[0], message_id: cm[1], text: ownerText_(o), parse_mode: 'HTML',
      reply_markup: { inline_keyboard: buttons_(o) }, disable_web_page_preview: true
    });
  });
}

function notifyBuyer_(o) {
  if (!o.buyerChat) return;
  const shop = settings_();
  const t = {
    work: 'Флорист взял ваш заказ <b>№' + o.no + '</b> в работу 🛠\nСкоро свяжемся, чтобы подтвердить детали и отправить ссылку на оплату.',
    ready: o.pickup
      ? 'Ваш букет собран 💐 и ждёт вас: ' + esc_(shop['Самовывоз'] || '') + ', ' + o.date + ', ' + o.interval + '.'
      : 'Ваш букет собран 💐' + (o.photo && !o.photoSent ? ' Фото пришлём перед отправкой.' : ' Скоро передадим курьеру.'),
    way: 'Курьер выехал 🚚 Заказ <b>№' + o.no + '</b> будет у получателя в интервале ' + o.interval + '.',
    done: 'Заказ <b>№' + o.no + '</b> доставлен ✅\nСпасибо, что выбрали ' + esc_(shop['Магазин'] || 'нас') + '! Будем рады видеть вас снова 🌸',
    given: 'Заказ <b>№' + o.no + '</b> выдан ✅\nСпасибо, что выбрали ' + esc_(shop['Магазин'] || 'нас') + '! 🌸',
    cancel: 'Заказ <b>№' + o.no + '</b> отменён. Если это ошибка — позвоните нам: ' + esc_(shop['Телефон магазина'] || '') + '.'
  }[o.code];
  if (t) say_(o.buyerChat, t);
}

function dayList_(chat, plus) {
  const d = new Date(Date.now() + plus * 864e5);
  const ru = Utilities.formatDate(d, TZ, 'dd.MM.yyyy');
  const list = ordersRows_().map(function (r, i) { return { r: r, row: i + 2 }; }).filter(function (x) {
    return d2s_(x.r[O['Дата']], 'dd.MM.yyyy') === ru && String(x.r[O['Статус']]) !== ST.cancel;
  }).map(function (x) { return orderAt_(x.row); });
  list.sort(function (a, b) { return a.slot - b.slot; });
  if (!list.length) return say_(chat, (plus ? 'На завтра' : 'На сегодня') + ' заказов нет.');
  const sum = list.reduce(function (s, o) { return s + o.total; }, 0);
  say_(chat, '<b>' + (plus ? 'Завтра' : 'Сегодня') + ', ' + ruDate_(ru) + '</b> — ' + list.length + ' зак. на ' + money_(sum) + '\n\n' +
    list.map(function (o) {
      return ST_ICON[o.code] + ' ' + o.interval + ' · №' + o.no + ' · ' + esc_(o.item) + ' ' + o.size + ' · ' + money_(o.total);
    }).join('\n'));
}

function ownerText_(o) {
  const L = [];
  L.push(ST_ICON[o.code] + ' <b>Заказ №' + o.no + '</b> · ' + money_(o.total));
  L.push('');
  L.push('💐 ' + esc_(o.item) + ', размер ' + o.size + ' — ' + money_(o.bouquet));
  L.push('🗓 ' + whenText_(o));
  L.push((o.pickup ? '🏠 ' : '📍 ') + esc_(o.address) + (o.pickup ? '' : ' — ' + (o.delivery ? money_(o.delivery) : 'бесплатно')));
  L.push('');
  L.push('👤 ' + esc_(o.name) + ', ' + esc_(o.phone));
  if (o.recipient && o.recipient !== 'Сам покупатель') L.push('🎁 Получатель: ' + esc_(o.recipient) + (o.rphone ? ', ' + esc_(o.rphone) : ''));
  if (o.card) L.push('✉️ Открытка: «' + esc_(o.card) + '»');
  if (o.anon) L.push('🕶 анонимно');
  if (o.photoSent) L.push('📷 Фото букета отправлено покупателю ✅');
  else if (o.photo && o.buyerChat && o.code !== 'cancel') L.push('📷 <b>Ждёт фото букета</b> — ответьте на это сообщение фотографией, бот перешлёт покупателю');
  else if (o.photo && o.code !== 'cancel') L.push('📷 Просит фото букета — покупатель без Telegram, отправьте в WhatsApp');
  if (o.comment) L.push('💬 ' + esc_(o.comment));
  L.push('');
  L.push('Статус: <b>' + o.status + '</b>' + (o.taker && o.code !== 'new' ? ' · взял(а) ' + esc_(o.taker) : ''));
  if (o.buyerChat) L.push('🔔 Покупатель получает статусы в Telegram');
  return L.join('\n');
}

function buttons_(o) {
  const label = {
    work: '🛠 Взять в работу', ready: '💐 Собран', way: '🚚 Передан курьеру',
    done: '✅ Доставлен', given: '✅ Выдан', cancel: '✖️ Отменить'
  };
  const nx = next_(o);
  if (!nx.length) return [];
  const main = nx.filter(function (c) { return c !== 'cancel'; }).map(function (c) {
    return { text: label[c], callback_data: 's:' + o.no + ':' + c };
  });
  const rows = main.length ? [main] : [];
  if (nx.indexOf('cancel') >= 0) rows.push([{ text: label.cancel, callback_data: 's:' + o.no + ':cancel' }]);
  return rows;
}

function next_(o) {
  return ({
    new: ['work', 'cancel'],
    work: ['ready', 'cancel'],
    ready: o.pickup ? ['given', 'cancel'] : ['way', 'cancel'],
    way: ['done'],
    done: [], given: [], cancel: []
  })[o.code] || [];
}

// ===== ДОСТУП К ДАННЫМ =====

function orderAt_(row) {
  const r = sheet_(SH.ord).getRange(row, 1, 1, ORD_H.length).getValues()[0];
  const status = String(r[O['Статус']]);
  const code = Object.keys(ST).filter(function (k) { return ST[k] === status; })[0] || 'new';
  let msgs = [];
  try { msgs = JSON.parse(String(r[O['тех: сообщения']] || '[]')); } catch (e) { msgs = []; }
  const tgCell = String(r[O['Telegram покупателя']] || '');
  return {
    row: row, no: Number(r[O['№']]), status: status, code: code,
    item: String(r[O['Букет']]), size: String(r[O['Размер']]),
    bouquet: Number(r[O['Букет, ₽']]) || 0, delivery: Number(r[O['Доставка, ₽']]) || 0, total: Number(r[O['Итого, ₽']]) || 0,
    date: d2s_(r[O['Дата']], 'dd.MM.yyyy'), interval: String(r[O['Интервал']]),
    slot: Number(String(r[O['Интервал']]).split(':')[0]) || 0,
    pickup: String(r[O['Получение']]) === 'Самовывоз', address: String(r[O['Адрес']]),
    name: String(r[O['Покупатель']]), phone: String(r[O['Телефон']]),
    recipient: String(r[O['Получатель']]), rphone: String(r[O['Телефон получателя']]),
    card: String(r[O['Открытка']]), photo: /^(Да|Отправлено)$/.test(String(r[O['Фото до отправки']])), photoSent: String(r[O['Фото до отправки']]) === 'Отправлено', anon: String(r[O['Анонимно']]) === 'Да',
    comment: String(r[O['Комментарий']]), taker: String(r[O['Взял в работу']]),
    buyerChat: tgCell ? tgCell.split(' ')[0] : '', msgs: msgs
  };
}

function ordersRows_() {
  const v = sheet_(SH.ord).getDataRange().getValues();
  return v.slice(1).filter(function (r) { return r[0] !== ''; });
}

function findOrder_(no) {
  const v = sheet_(SH.ord).getDataRange().getValues();
  for (let i = 1; i < v.length; i++) if (Number(v[i][0]) === no) return i + 1;
  return 0;
}

function sheet_(name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) throw new Error('Нет листа «' + name + '». Запустите setup().');
  return sh;
}

function props_() { return PropertiesService.getScriptProperties(); }
function botName_() { return props_().getProperty('BOT') || ''; }
function owners_() { return JSON.parse(props_().getProperty('OWNERS') || '[]'); }
function addOwner_(chat) {
  const o = owners_();
  if (o.indexOf(String(chat)) < 0) o.push(String(chat));
  props_().setProperty('OWNERS', JSON.stringify(o));
}

// ===== УТИЛИТЫ =====

function tg_(method, payload) {
  try {
    const r = UrlFetchApp.fetch('https://api.telegram.org/bot' + BOT_TOKEN + '/' + method, {
      method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true
    });
    return JSON.parse(r.getContentText());
  } catch (e) {
    console.error(e);
    return null;
  }
}

function say_(chat, text) {
  return tg_('sendMessage', { chat_id: chat, text: text, parse_mode: 'HTML', disable_web_page_preview: true });
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function whenText_(o) { return ruDate_(o.date) + ', ' + o.interval; }

function ruDate_(dmy) {
  const m = String(dmy).match(/(\d{2})\.(\d{2})\.(\d{4})/);
  if (!m) return String(dmy);
  const d = new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
  return Number(m[1]) + ' ' + MONTHS[Number(m[2]) - 1] + ', ' + WD[d.getDay()];
}

function isDate_(v) { return Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v); }
function d2s_(v, f) { return isDate_(v) ? Utilities.formatDate(v, TZ, f) : String(v == null ? '' : v); }
/** «05.11.2026» или «05.11.2026 13:06» → Date (дата без времени — на полдень, чтобы не уехать на соседний день). */
function parseRu_(v) {
  if (isDate_(v)) return v;
  const m = String(v || '').match(/(\d{2})\.(\d{2})\.(\d{4})(?:\s+(\d{1,2}):(\d{2}))?/);
  if (!m) return v;
  return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]), m[4] ? Number(m[4]) : 12, m[5] ? Number(m[5]) : 0);
}
function isoToday_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function nowStr_() { return Utilities.formatDate(new Date(), TZ, 'dd.MM.yyyy HH:mm'); }
function r100_(n) { return Math.round(n / 100) * 100; }
function money_(n) { return String(Math.round(n)).replace(/\B(?=(\d{3})+(?!\d))/g, ' ') + ' ₽'; }
function esc_(s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }
function clean_(s, max) { return String(s == null ? '' : s).replace(/\s+/g, ' ').trim().slice(0, max); }
/** Текст в ячейку без превращения в формулу или число. */
function txt_(s) { s = String(s == null ? '' : s); return s === '' ? '' : "'" + s; }
function personName_(u) { return u ? [u.first_name, u.last_name].filter(Boolean).join(' ') || ('@' + u.username) : ''; }

function phoneFmt_(p) {
  let d = String(p || '').replace(/\D/g, '');
  if (d.length === 11 && d[0] === '8') d = '7' + d.slice(1);
  if (d.length !== 11 || d[0] !== '7') return '';
  return '+7 ' + d.slice(1, 4) + ' ' + d.slice(4, 7) + '-' + d.slice(7, 9) + '-' + d.slice(9, 11);
}

// ===== ОФОРМЛЕНИЕ И СВОДКА =====

const C_MOSS = '#3B4213', C_CREAM = '#FFF8EA', C_BAND = '#F7F0DF', C_RUST = '#9C4A1E', C_INK = '#2A2F0C';
const F_MONEY = '#,##0 "₽"', F_DATE = 'dd.mm.yyyy', F_DT = 'dd.mm.yyyy hh:mm';
const ST_COLORS = {
  'Новый': ['#FFF1C2', '#7A5B00'], 'В работе': ['#DCEBFF', '#1E4E8C'], 'Собран': ['#EDE3FF', '#5B3A9E'],
  'В пути': ['#EDE3FF', '#5B3A9E'], 'Доставлен': ['#DDF2DF', '#1F6B2A'], 'Выдан': ['#DDF2DF', '#1F6B2A'],
  'Отменён': ['#ECECEC', '#7A7A7A']
};

/**
 * Обновление уже работающей таблицы: переводит старые даты, оформляет листы, строит «Сводку».
 * Запускать после замены кода. Данные не удаляются. Можно запускать повторно.
 */
function upgrade() {
  const ss = SpreadsheetApp.getActive();
  Logger.log('Версия кода: 3 (автоподбор разделителя формул)');
  Logger.log('Язык таблицы: ' + ss.getSpreadsheetLocale());
  ss.setSpreadsheetTimeZone(TZ);
  const steps = [['Перевод дат', migrateDates_], ['Оформление листов', styleAll_], ['Сводка', buildSummary_]];
  let fails = 0;
  steps.forEach(function (st) {
    try { st[1](); Logger.log('✅ ' + st[0]); }
    catch (e) { fails++; Logger.log('❌ ' + st[0] + ': ' + (e && e.message ? e.message : e)); }
  });
  Logger.log(fails ? 'Есть ошибки — пришлите скриншот журнала.' : 'Готово: таблица оформлена, «Сводка» построена.');
}

/** Какой разделитель аргументов понимает эта таблица: «,» или «;». */
function formulaSep_(sh) {
  const c = sh.getRange('Z1');
  c.setFormula('=SUM(1,2)');
  SpreadsheetApp.flush();
  const ok = c.getValue() === 3;
  c.clearContent();
  return ok ? ',' : ';';
}

/** Заменяет запятые-разделители на «;» вне строк в кавычках. */
function toSep_(f, sep) {
  if (sep === ',') return f;
  let out = '', q = false;
  for (let i = 0; i < f.length; i++) {
    const ch = f.charAt(i);
    if (ch === '"') q = !q;
    out += (!q && ch === ',') ? sep : ch;
  }
  return out;
}

function migrateDates_() {
  const ord = sheet_(SH.ord), n = ord.getLastRow() - 1;
  if (n > 0) {
    [O['Создан'], O['Дата'], O['Обновлён']].forEach(function (ci) {
      const rg = ord.getRange(2, ci + 1, n, 1);
      rg.setValues(rg.getValues().map(function (r) { return [parseRu_(r[0])]; }));
    });
  }
  const cli = sheet_(SH.cli), m = cli.getLastRow() - 1;
  if (m > 0) {
    const rg = cli.getRange(2, 5, m, 2);
    rg.setValues(rg.getValues().map(function (r) { return [parseRu_(r[0]), parseRu_(r[1])]; }));
  }
}

function styleSheet_(sh, widths, tab, rowH) {
  const cols = widths.length, rows = Math.max(sh.getMaxRows(), 500);
  if (sh.getMaxRows() < rows) sh.insertRowsAfter(sh.getMaxRows(), rows - sh.getMaxRows());
  sh.setHiddenGridlines(true);
  sh.setTabColor(tab);
  sh.getBandings().forEach(function (b) { b.remove(); });
  const all = sh.getRange(1, 1, rows, cols);
  all.setFontFamily('Roboto').setFontSize(11).setFontColor(C_INK).setVerticalAlignment('middle');
  all.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false)
    .setHeaderRowColor(C_MOSS).setFirstRowColor('#FFFFFF').setSecondRowColor(C_BAND);
  sh.getRange(1, 1, 1, cols).setBackground(C_MOSS).setFontWeight('bold').setFontColor(C_CREAM).setFontSize(11)
    .setWrap(true).setHorizontalAlignment('center');
  sh.setRowHeight(1, 44);
  if (rowH) sh.setRowHeights(2, rows - 1, rowH);
  sh.setFrozenRows(1);
  widths.forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
}

function fmtCol_(sh, col, f, align) {
  const rg = sh.getRange(2, col, sh.getMaxRows() - 1, 1);
  if (f) rg.setNumberFormat(f);
  if (align) rg.setHorizontalAlignment(align);
  return rg;
}

function styleAll_() {
  const ss = SpreadsheetApp.getActive();

  // Заказы
  const ord = sheet_(SH.ord);
  styleSheet_(ord, [70, 140, 120, 210, 75, 110, 110, 120, 110, 120, 110, 300, 170, 170, 160, 170, 220, 110, 100, 260, 150, 140, 170, 60, 60], C_RUST, 0);
  const oc = function (h) { return O[h] + 1; };
  fmtCol_(ord, oc('№'), '0', 'center');
  fmtCol_(ord, oc('Создан'), F_DT, 'center');
  fmtCol_(ord, oc('Статус'), null, 'center').setFontWeight('bold');
  fmtCol_(ord, oc('Размер'), null, 'center');
  ['Букет, ₽', 'Доставка, ₽', 'Итого, ₽'].forEach(function (h) { fmtCol_(ord, oc(h), F_MONEY, 'right'); });
  fmtCol_(ord, oc('Итого, ₽'), F_MONEY, 'right').setFontWeight('bold');
  fmtCol_(ord, oc('Дата'), F_DATE, 'center');
  fmtCol_(ord, oc('Интервал'), null, 'center');
  fmtCol_(ord, oc('Получение'), null, 'center');
  fmtCol_(ord, oc('Фото до отправки'), null, 'center');
  fmtCol_(ord, oc('Анонимно'), null, 'center');
  fmtCol_(ord, oc('Обновлён'), F_DT, 'center');
  ['Адрес', 'Открытка', 'Комментарий'].forEach(function (h) {
    fmtCol_(ord, oc(h)).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  });
  ord.hideColumns(oc('тех: сообщения'), 2);
  ord.setFrozenColumns(1);
  const stRange = ord.getRange(2, oc('Статус'), ord.getMaxRows() - 1, 1);
  const rules = Object.keys(ST_COLORS).map(function (s) {
    return SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(s)
      .setBackground(ST_COLORS[s][0]).setFontColor(ST_COLORS[s][1]).setRanges([stRange]).build();
  });
  ord.setConditionalFormatRules(rules);

  // Клиенты
  const cli = sheet_(SH.cli);
  styleSheet_(cli, [180, 200, 100, 130, 140, 150, 240, 160], C_MOSS, 30);
  fmtCol_(cli, 3, '0', 'center');
  fmtCol_(cli, 4, F_MONEY, 'right').setFontWeight('bold');
  fmtCol_(cli, 5, F_DATE, 'center');
  fmtCol_(cli, 6, F_DATE, 'center');
  cli.setConditionalFormatRules([SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(1)
    .setBackground('#DDF2DF').setFontColor('#1F6B2A').setRanges([cli.getRange(2, 3, cli.getMaxRows() - 1, 1)]).build()]);

  // Каталог
  const cat = sheet_(SH.cat);
  styleSheet_(cat, [60, 240, 140, 170, 320, 170, 170, 120], '#EE8B6A', 34);
  fmtCol_(cat, 1, '0', 'center');
  fmtCol_(cat, 3, null, 'center').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(Object.keys(CATS), true).setAllowInvalid(false).build());
  fmtCol_(cat, 4, F_MONEY, 'right').setFontWeight('bold');
  fmtCol_(cat, 5).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  const stock = fmtCol_(cat, 8, null, 'center');
  stock.setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['Да', 'Нет'], true).setAllowInvalid(false).build());
  cat.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Нет').setBackground('#F6D9D0').setFontColor(C_RUST).setRanges([stock]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Да').setBackground('#DDF2DF').setFontColor('#1F6B2A').setRanges([stock]).build()
  ]);

  // Настройки
  const set = sheet_(SH.set);
  styleSheet_(set, [280, 560], '#9AA0A6', 30);
  set.getRange(2, 2, set.getMaxRows() - 1, 1).setHorizontalAlignment('left');

  const def = ss.getSheetByName('Лист1') || ss.getSheetByName('Sheet1');
  if (def && def.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(def);
}

function buildSummary_() {
  const ss = SpreadsheetApp.getActive();
  const NAME = 'Сводка';
  const old = ss.getSheetByName(NAME);
  if (old) ss.deleteSheet(old);
  const sh = ss.insertSheet(NAME, 0);
  sh.setHiddenGridlines(true);
  const SEP = formulaSep_(sh);
  Logger.log('Разделитель формул в этой таблице: «' + SEP + '»');
  sh.setTabColor(C_MOSS);
  const shop = settings_()['Магазин'] || 'Магазин';
  const Z = "'" + SH.ord + "'!", K = "'" + SH.cli + "'!";
  const base = Z + 'A:A,">0",' + Z + 'C:C,"<>' + ST.cancel + '"';
  const P = [
    ['Сегодня', ',' + Z + 'B:B,">="&TODAY()'],
    ['7 дней', ',' + Z + 'B:B,">="&TODAY()-6'],
    ['Этот месяц', ',' + Z + 'B:B,">="&EOMONTH(TODAY(),-1)+1'],
    ['Всё время', '']
  ];
  const col = function (i) { return String.fromCharCode(66 + i); };   // B, C, D, E
  // Формулы пишем через setFormula: он всегда понимает английский синтаксис с запятыми,
  // а setValues разбирает формулы по языку таблицы (в русской — через «;») и даёт #ERROR!.
  const put = function (a1, rows) {
    const r = sh.getRange(a1).offset(0, 0, rows.length, rows[0].length);
    rows.forEach(function (row, i) {
      row.forEach(function (v, j) {
        const cell = r.offset(i, j, 1, 1);
        if (typeof v === 'string' && v.charAt(0) === '=') cell.setFormula(toSep_(v, SEP)); else cell.setValue(v);
      });
    });
    return r;
  };
  const head = function (a1, cols) {
    const r = put(a1, [cols]);
    r.setBackground(C_MOSS).setFontColor(C_CREAM).setFontWeight('bold').setFontSize(11);
    r.offset(0, 1, 1, cols.length - 1).setHorizontalAlignment('right');
    return r;
  };

  sh.getRange(1, 1, 120, 12).setFontFamily('Roboto').setFontSize(11).setFontColor(C_INK).setVerticalAlignment('middle');
  sh.setColumnWidth(1, 290);
  [2, 3, 4, 5].forEach(function (c) { sh.setColumnWidth(c, 135); });
  sh.setColumnWidth(6, 30);
  [7, 8, 9, 10].forEach(function (c) { sh.setColumnWidth(c, 130); });
  sh.setColumnWidth(7, 230);

  put('A1', [[shop.toUpperCase() + ' · Сводка']]).setFontFamily('Oranienbaum').setFontSize(30).setFontColor(C_MOSS);
  sh.setRowHeight(1, 56);
  put('A2', [['Управленческая отчётность · считается сама при каждом заказе · отменённые заказы не учитываются']])
    .setFontColor('#7A7F5C').setFontSize(10);

  // 1. Деньги и заказы
  head('A4', ['Деньги и заказы'].concat(P.map(function (p) { return p[0]; })));
  put('A5', [['Выручка'].concat(P.map(function (p) { return '=SUMIFS(' + Z + 'H:H,' + base + p[1] + ')'; }))]);
  put('A6', [['Заказов'].concat(P.map(function (p) { return '=COUNTIFS(' + base + p[1] + ')'; }))]);
  put('A7', [['Средний чек'].concat(P.map(function (p, i) { return '=IFERROR(' + col(i) + '5/' + col(i) + '6,0)'; }))]);
  put('A8', [['Отменено заказов'].concat(P.map(function (p) {
    return '=COUNTIFS(' + Z + 'A:A,">0",' + Z + 'C:C,"' + ST.cancel + '"' + p[1] + ')';
  }))]);
  sh.getRange('B5:E5').setNumberFormat(F_MONEY).setFontSize(16).setFontWeight('bold').setFontColor(C_RUST);
  sh.getRange('B7:E7').setNumberFormat(F_MONEY);
  sh.getRange('B6:E6').setNumberFormat('0');
  sh.getRange('B8:E8').setNumberFormat('0').setFontColor('#7A7A7A');
  sh.getRange('B5:E8').setHorizontalAlignment('right');
  sh.setRowHeight(5, 38);

  // 2. В работе сейчас
  head('A10', ['Заказы в работе сейчас', 'Заказов', 'На сумму']);
  const live = [['🆕 Новые — ждут ответа', ST.new], ['🛠 В работе', ST.work], ['💐 Собраны', ST.ready], ['🚚 В пути', ST.way]];
  put('A11', live.map(function (x) {
    return [x[0], '=COUNTIFS(' + Z + 'C:C,"' + x[1] + '")', '=SUMIFS(' + Z + 'H:H,' + Z + 'C:C,"' + x[1] + '")'];
  }));
  sh.getRange('C11:C14').setNumberFormat(F_MONEY);
  sh.getRange('B11:C14').setHorizontalAlignment('right');
  sh.setConditionalFormatRules([SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0)
    .setBackground('#FFF1C2').setFontColor('#7A5B00').setBold(true).setRanges([sh.getRange('B11')]).build()]);

  // 3. Клиенты
  head('A16', ['Клиенты', '']);
  put('A17', [
    ['Всего в базе', '=COUNTA(' + K + 'A2:A)'],
    ['Новых за этот месяц', '=COUNTIFS(' + K + 'E2:E,">="&EOMONTH(TODAY(),-1)+1)'],
    ['Повторных (2+ заказа)', '=COUNTIF(' + K + 'C2:C,">1")'],
    ['Доля повторных', '=IFERROR(B19/B17,0)'],
    ['Выручка на одного клиента', '=IFERROR(SUM(' + K + 'D2:D)/B17,0)']
  ]);
  sh.getRange('B20').setNumberFormat('0%');
  sh.getRange('B21').setNumberFormat(F_MONEY);
  sh.getRange('B17:B19').setNumberFormat('0');
  sh.getRange('B17:B21').setHorizontalAlignment('right');

  // 4. Топ букетов
  head('A23', ['Топ букетов', 'Штук', 'Выручка']);
  put('A24', [['=IFERROR(QUERY(' + Z + "A2:H,\"select D, count(A), sum(H) where A is not null and C <> '" + ST.cancel +
    "' group by D order by sum(H) desc limit 5 label count(A) '', sum(H) ''\",0),\"Пока нет заказов\")"]]);
  sh.getRange('B24:B28').setNumberFormat('0').setHorizontalAlignment('right');
  sh.getRange('C24:C28').setNumberFormat(F_MONEY).setHorizontalAlignment('right');

  // 5. Как покупают
  const cnt = function (cond) { return '=COUNTIFS(' + base + ',' + cond + ')'; };
  const sum = function (cond) { return '=SUMIFS(' + Z + 'H:H,' + base + ',' + cond + ')'; };
  const rows = [];
  rows.push(['Курьер', cnt(Z + 'K:K,"Курьер"'), sum(Z + 'K:K,"Курьер"')]);
  zones_(settings_()).forEach(function (z) {
    rows.push(['   · ' + z.n, cnt(Z + 'L:L,"*(' + z.n + ')*"'), sum(Z + 'L:L,"*(' + z.n + ')*"')]);
  });
  rows.push(['Самовывоз', cnt(Z + 'K:K,"Самовывоз"'), sum(Z + 'K:K,"Самовывоз"')]);
  rows.push(['', '', '']);
  SLOTS.forEach(function (s) {
    const t = s[0] + ':00–' + s[1] + ':00';
    rows.push(['Интервал ' + t, cnt(Z + 'J:J,"' + t + '"'), sum(Z + 'J:J,"' + t + '"')]);
  });
  rows.push(['', '', '']);
  ['S', 'M', 'L'].forEach(function (z) { rows.push(['Размер ' + z, cnt(Z + 'E:E,"' + z + '"'), sum(Z + 'E:E,"' + z + '"')]); });
  head('A30', ['Как покупают', 'Заказов', 'Выручка']);
  put('A31', rows);
  sh.getRange(31, 2, rows.length, 1).setNumberFormat('0').setHorizontalAlignment('right');
  sh.getRange(31, 3, rows.length, 1).setNumberFormat(F_MONEY).setHorizontalAlignment('right');

  // 6. Праздники (справа под графиком)
  const H = [
    ['Новый год', 12, 31], ['14 февраля', 2, 14], ['8 Марта', 3, 8],
    ['День знаний', 9, 1], ['День учителя', 10, 5], ['День матери', 0, 0]
  ];
  const nextDate = function (m, d) {
    if (!m) {   // последнее воскресенье ноября
      const f = function (y) { return 'DATE(' + y + ',11,30)-MOD(WEEKDAY(DATE(' + y + ',11,30))-1,7)'; };
      return '=IF(' + f('YEAR(TODAY())') + '>=TODAY(),' + f('YEAR(TODAY())') + ',' + f('YEAR(TODAY())+1') + ')';
    }
    return '=IF(DATE(YEAR(TODAY()),' + m + ',' + d + ')>=TODAY(),DATE(YEAR(TODAY()),' + m + ',' + d + '),DATE(YEAR(TODAY())+1,' + m + ',' + d + '))';
  };
  put('N4', [['Праздник', 'Дата']]);
  put('N5', H.map(function (h) { return [h[0], nextDate(h[1], h[2])]; }));
  head('G23', ['Ближайшие праздники', 'Дата', 'Дней', 'Клиентов в базе']);
  put('G24', [['=SORT(N5:O10,2,TRUE)']]);
  put('I24', H.map(function (h, i) { return ['=H' + (24 + i) + '-TODAY()', '=B17']; }));
  put('G30', [['="🔔 Ближайший — "&G24&" через "&I24&" дн. "&IF(I24<=14,"Пора напомнить клиентам из базы.","Напоминать за 7–14 дней.")']])
    .setFontColor(C_RUST).setFontWeight('bold');
  sh.getRange('H24:H29').setNumberFormat('dd.mm.yyyy').setHorizontalAlignment('right');
  sh.getRange('I24:J29').setNumberFormat('0').setHorizontalAlignment('right');
  sh.setConditionalFormatRules(sh.getConditionalFormatRules().concat([
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$I24<=14')
      .setBackground('#FFF1C2').setFontColor('#7A5B00').setRanges([sh.getRange('G24:J29')]).build()
  ]));

  // 7. График выручки по дням (данные справа, в N:O)
  put('N13', [['День', 'Выручка']]);
  put('N14', [['=SEQUENCE(30,1,TODAY()-29)']]);
  put('O14', [['=MAP(N14:N43,LAMBDA(d,SUMIFS(' + Z + 'H:H,' + base + ',' + Z + 'B:B,">="&d,' + Z + 'B:B,"<"&(d+1))))']]);
  sh.getRange('N14:N43').setNumberFormat('dd.mm');
  sh.getRange('O14:O43').setNumberFormat(F_MONEY);
  sh.getRange('N3').setValue('данные для графика и праздников').setFontColor('#9AA0A6').setFontSize(9);
  sh.getRange('N4:O43').setFontColor('#9AA0A6').setFontSize(9);
  const chart = sh.newChart().setChartType(Charts.ChartType.COLUMN)
    .addRange(sh.getRange('N13:O43'))
    .setNumHeaders(1)
    .setPosition(4, 7, 0, 0)
    .setOption('title', 'Выручка по дням, последние 30 дней')
    .setOption('titleTextStyle', { color: C_INK, fontSize: 13, bold: true })
    .setOption('legend', { position: 'none' })
    .setOption('colors', [C_RUST])
    .setOption('backgroundColor', '#FFFFFF')
    .setOption('vAxis', { format: '#,##0 ₽', gridlines: { color: '#EEEEEE' } })
    .setOption('hAxis', { format: 'dd.MM', textStyle: { fontSize: 9 } })
    .setOption('width', 640).setOption('height', 340)
    .build();
  sh.insertChart(chart);

  sh.setFrozenRows(2);
  SpreadsheetApp.flush();
  const bad = sh.getRange('B5').getDisplayValue();
  Logger.log('Проверка B5 (выручка сегодня): ' + bad);
  ss.setActiveSheet(sh);
  return sh;
}

// ===== ОФОРМЛЕНИЕ БОТА =====

/** Имя, описание, кнопка «Каталог» и команды бота. Запускать после изменений, можно повторно. */
function styleBot() {
  const st = settings_(), shop = st['Магазин'] || 'Pivoine';
  const log = function (m, r) { Logger.log((r && r.ok ? '✅ ' : '❌ ') + m + (r && r.ok ? '' : ': ' + JSON.stringify(r))); };
  log('Имя', tg_('setMyName', { name: shop + ' · цветочная мастерская' }));
  const shortD = 'Цветочная мастерская ' + shop + ' 🌸 Авторские букеты с доставкой по Москве за 2 часа.';
  const longD = 'Вас приветствует цветочная мастерская ' + shop + ' 🌸\n\n' +
    'Авторские букеты с доставкой по Москве за 2 часа, ежедневно с 9:00 до 21:00.\n\n' +
    'Здесь вы можете:\n💐 выбрать букет в каталоге и оформить заказ за минуту\n' +
    '📷 получить фото букета перед отправкой\n🔔 следить за заказом: собран, передан курьеру, доставлен\n\n' +
    'Нажмите «Начать», чтобы открыть каталог.';
  ['', 'ru'].forEach(function (lc) {        // по умолчанию и отдельно для русского интерфейса
    const x = lc ? { language_code: lc } : {};
    log('Короткое описание ' + (lc || 'общее'), tg_('setMyShortDescription', Object.assign({ short_description: shortD }, x)));
    log('Описание ' + (lc || 'общее'), tg_('setMyDescription', Object.assign({ description: longD }, x)));
  });
  log('Кнопка «Каталог»', tg_('setChatMenuButton', {
    menu_button: { type: 'web_app', text: 'Каталог', web_app: { url: SITE_URL } }
  }));
  log('Команды покупателя', tg_('setMyCommands', {
    commands: [{ command: 'catalog', description: 'Каталог букетов' }, { command: 'start', description: 'О магазине' }]
  }));
  owners_().forEach(function (chat) {
    log('Команды владельца', tg_('setMyCommands', {
      scope: { type: 'chat', chat_id: Number(chat) },
      commands: [
        { command: 'today', description: 'Заказы на сегодня' },
        { command: 'tomorrow', description: 'Заказы на завтра' },
        { command: 'catalog', description: 'Открыть сайт' }
      ]
    }));
  });
  Logger.log('Аватарку бота ставим вручную: @BotFather → /setuserpic → выбрать бота → отправить картинку.');
}

function catalogButton_() {
  return { inline_keyboard: [[{ text: '🌸 Открыть каталог', web_app: { url: SITE_URL } }]] };
}

// ===== ДЕМО-ДАННЫЕ =====

/**
 * Заполняет «Заказы» и «Клиенты» правдоподобными заказами за последние 30 дней — для показа клиентам.
 * ВНИМАНИЕ: перед заполнением очищает эти два листа. Каталог и настройки не трогает.
 */
function seedDemo() {
  let seed = 20261105;
  const rnd = function () { seed |= 0; seed = seed + 0x6D2B79F5 | 0; let t = Math.imul(seed ^ seed >>> 15, 1 | seed); t = t + Math.imul(t ^ t >>> 7, 61 | t) ^ t; return ((t ^ t >>> 14) >>> 0) / 4294967296; };
  const pick = function (a) { return a[Math.floor(rnd() * a.length)]; };
  const wpick = function (a) { const s = a.reduce(function (x, y) { return x + y[1]; }, 0); let r = rnd() * s; for (let i = 0; i < a.length; i++) { r -= a[i][1]; if (r <= 0) return a[i][0]; } return a[0][0]; };

  const people = ['Анна Смирнова', 'Мария Кузнецова', 'Елена Попова', 'Ольга Волкова', 'Дмитрий Соколов', 'Алексей Морозов',
    'Сергей Новиков', 'Ирина Лебедева', 'Наталья Козлова', 'Татьяна Павлова', 'Андрей Семёнов', 'Екатерина Голубева',
    'Максим Виноградов', 'Юлия Богданова', 'Павел Фёдоров', 'Светлана Михайлова', 'Артём Белов', 'Ксения Тарасова',
    'Николай Орлов', 'Виктория Киселёва', 'Игорь Макаров', 'Полина Андреева', 'Дарья Захарова', 'Роман Никитин',
    'Алина Громова', 'Евгений Зайцев', 'Вера Комарова', 'Кирилл Борисов', 'Марина Егорова', 'Денис Ковалёв',
    'Софья Медведева', 'Олег Степанов', 'Людмила Яковлева', 'Илья Гусев', 'Анастасия Филиппова', 'Георгий Власов'].map(function (n, i) {
    const d = String(9150000000 + Math.floor(rnd() * 849999999));
    return { name: n, phone: phoneFmt_('7' + d), tg: rnd() < 0.45 ? '@' + translit_(n.split(' ')[0]).toLowerCase() + '_' + (10 + i) : '' };
  });
  const loyal = people.slice(0, 6);   // эти заказывают по 2–3 раза
  const streets = ['ул. Тверская, 18', 'Кутузовский пр-т, 24', 'ул. Покровка, 31', 'Ленинский пр-т, 62', 'ул. Профсоюзная, 45',
    'Мичуринский пр-т, 9', 'ул. Большая Дмитровка, 7', 'Ломоносовский пр-т, 14', 'ул. Маросейка, 11', 'пр-т Мира, 102',
    'ул. Остоженка, 25', 'Хорошёвское ш., 38', 'ул. Арбат, 44', 'ул. Пятницкая, 56', 'Варшавское ш., 87'];
  const far = ['Красногорск, ул. Ленина, 12', 'Химки, ул. Молодёжная, 4', 'Мытищи, Новомытищинский пр-т, 30', 'Одинцово, ул. Неделина, 8'];
  const cards = ['С днём рождения, мама! Люблю тебя ❤️', 'Спасибо, что ты есть', 'С годовщиной, любимая!', 'Поздравляю с новой должностью!',
    'Просто так 🌸', 'Выздоравливай скорее!', 'С днём рождения! Будь счастлива', 'Прости меня'];
  const comments = ['Позвонить получателю за час', 'Домофон не работает, звонить на телефон', 'Сюрприз — не говорить от кого',
    'Оставить у консьержа', 'Можно чуть больше зелени', 'Подъезд со двора'];
  const recips = ['Мама', 'Ольга', 'Алина', 'Светлана Петровна', 'Катя', 'Наталья'];
  const takers = ['Мария', 'Арсений'];

  const st = settings_(), zones = zones_(st), cat = catalog_();
  const today = new Date(); today.setHours(12, 0, 0, 0);
  const orders = [];
  for (let back = 29; back >= 0; back--) {
    const day = new Date(today.getTime() - back * 864e5);
    const wd = day.getDay(), weekend = wd === 0 || wd === 5 || wd === 6;
    let n = Math.floor(rnd() * (weekend ? 4 : 3)) + (weekend ? 1 : 0);
    if (back === 0) n = 4;
    for (let k = 0; k < n; k++) {
      const p = rnd() < 0.2 ? pick(loyal) : pick(people.slice(6));
      const item = wpick(cat.map(function (c) { return [c, c.p < 10000 ? 3 : c.p < 15000 ? 2 : 1]; }));
      const size = wpick([['M', 6], ['S', 2], ['L', 2]]);
      const pickup = rnd() < 0.2;
      const zi = pickup ? -1 : wpick([[0, 6], [1, 3], [2, 1]]);
      const zone = zi >= 0 ? zones[zi] || zones[0] : null;
      const nowH = new Date().getHours();
      const hour = back === 0 ? 9 + Math.floor(rnd() * Math.max(1, Math.min(12, nowH - 9))) : 9 + Math.floor(rnd() * 12);
      const created = new Date(day.getTime()); created.setHours(hour, Math.floor(rnd() * 60), 0, 0);
      const plus = back === 0 ? (k < 2 ? 0 : 1) : (rnd() < 0.6 ? 0 : 1);
      const dDate = new Date(day.getTime() + plus * 864e5); dDate.setHours(12, 0, 0, 0);
      const okSlots = SLOTS.filter(function (s) { return plus || s[1] - 1 > hour + 2; });
      const slot = okSlots.length ? wpick(okSlots.map(function (s) { return [s, s[0] === 12 || s[0] === 15 ? 3 : 2]; })) : SLOTS[3];
      const other = rnd() < 0.55;
      orders.push({
        created: created, status: '', item: item, size: size, pickup: pickup, zone: zone,
        bouquet: r100_(item.p * SIZES[size]), delivery: pickup ? 0 : zone.p,
        date: dDate, slot: slot, p: p, back: back, plus: plus,
        street: pickup ? '' : (zi === 2 ? pick(far) : pick(streets)) + ', кв. ' + (1 + Math.floor(rnd() * 180)),
        rname: other ? pick(recips) : '', rphone: other ? phoneFmt_('79' + String(100000000 + Math.floor(rnd() * 899999999))) : '',
        card: rnd() < 0.4 ? pick(cards) : '', photo: rnd() < 0.6, anon: other && rnd() < 0.15,
        comment: rnd() < 0.3 ? pick(comments) : '', taker: pick(takers)
      });
    }
  }
  orders.sort(function (a, b) { return a.created - b.created; });
  // статусы: прошлое — доставлено, сегодняшние — по воронке, пара отмен
  const todays = orders.filter(function (o) { return o.back === 0; });
  orders.forEach(function (o) { o.status = o.pickup ? ST.given : ST.done; });
  orders.filter(function (o) { return o.back === 1 && o.plus === 1; }).forEach(function (o) { o.status = o.pickup ? ST.ready : ST.way; });
  [ST.new, ST.work, ST.ready, ST.way].forEach(function (s, i) {
    if (todays[i]) todays[i].status = (s === ST.way && todays[i].pickup) ? ST.ready : s;
  });
  [5, 17].forEach(function (i) { if (orders[i]) orders[i].status = ST.cancel; });

  const rows = orders.map(function (o, i) {
    const r = [];
    r[O['№']] = 1001 + i; r[O['Создан']] = o.created; r[O['Статус']] = o.status;
    r[O['Букет']] = txt_(o.item.n); r[O['Размер']] = o.size;
    r[O['Букет, ₽']] = o.bouquet; r[O['Доставка, ₽']] = o.delivery; r[O['Итого, ₽']] = o.bouquet + o.delivery;
    r[O['Дата']] = o.date; r[O['Интервал']] = txt_(o.slot[0] + ':00–' + o.slot[1] + ':00');
    r[O['Получение']] = o.pickup ? 'Самовывоз' : 'Курьер';
    r[O['Адрес']] = txt_(o.pickup ? 'Самовывоз: ' + (st['Самовывоз'] || '') : o.street + ' (' + o.zone.n + ')');
    r[O['Покупатель']] = txt_(o.p.name); r[O['Телефон']] = txt_(o.p.phone);
    r[O['Получатель']] = txt_(o.rname || 'Сам покупатель'); r[O['Телефон получателя']] = txt_(o.rphone);
    r[O['Открытка']] = txt_(o.card); r[O['Фото до отправки']] = o.photo ? 'Да' : 'Нет'; r[O['Анонимно']] = o.anon ? 'Да' : 'Нет';
    r[O['Комментарий']] = txt_(o.comment); r[O['Взял в работу']] = o.status === ST.new ? '' : txt_(o.taker);
    r[O['Обновлён']] = o.created; r[O['Telegram покупателя']] = '';
    r[O['тех: сообщения']] = '[]'; r[O['тех: ключ']] = 'demo';
    return r;
  });

  const cl = {};
  orders.forEach(function (o) {
    if (o.status === ST.cancel) return;
    const c = cl[o.p.phone] || (cl[o.p.phone] = { p: o.p, n: 0, sum: 0, first: o.created, last: o.created, item: '' });
    c.n++; c.sum += o.bouquet + o.delivery; c.last = o.created; c.item = o.item.n + ', ' + o.size;
  });
  const crow = Object.keys(cl).map(function (k) { return cl[k]; }).sort(function (a, b) { return b.sum - a.sum; }).map(function (c) {
    return [txt_(c.p.phone), txt_(c.p.name), c.n, c.sum, c.first, c.last, txt_(c.item), txt_(c.p.tg)];
  });

  clearSheet_(SH.ord); clearSheet_(SH.cli);
  sheet_(SH.ord).getRange(2, 1, rows.length, ORD_H.length).setValues(rows);
  sheet_(SH.cli).getRange(2, 1, crow.length, CLI_H.length).setValues(crow);
  SpreadsheetApp.flush();
  Logger.log('Готово: ' + rows.length + ' заказов, ' + crow.length + ' клиентов за последние 30 дней.');
}

/** Очищает «Заказы» и «Клиенты» (например, перед запуском у реального клиента). Каталог и настройки не трогает. */
function clearOrders() {
  clearSheet_(SH.ord); clearSheet_(SH.cli);
  Logger.log('Заказы и клиенты очищены.');
}

function clearSheet_(name) {
  const sh = sheet_(name), n = sh.getLastRow();
  if (n > 1) sh.getRange(2, 1, n - 1, sh.getLastColumn()).clearContent();
}

function translit_(s) {
  const m = { а: 'a', б: 'b', в: 'v', г: 'g', д: 'd', е: 'e', ё: 'e', ж: 'zh', з: 'z', и: 'i', й: 'y', к: 'k', л: 'l', м: 'm', н: 'n', о: 'o',
    п: 'p', р: 'r', с: 's', т: 't', у: 'u', ф: 'f', х: 'h', ц: 'ts', ч: 'ch', ш: 'sh', щ: 'sch', ы: 'y', э: 'e', ю: 'yu', я: 'ya', ь: '', ъ: '' };
  return String(s).toLowerCase().split('').map(function (c) { return m[c] !== undefined ? m[c] : c; }).join('');
}
