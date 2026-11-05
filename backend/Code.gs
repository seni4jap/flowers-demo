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
  const def = ss.getSheetByName('Лист1') || ss.getSheetByName('Sheet1');
  if (def && ss.getSheets().length > 1 && def.getLastRow() === 0) ss.deleteSheet(def);

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
  const d = String(r[O['Дата']]).match(/(\d{2})\.(\d{2})\.(\d{4})/);
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
  const dateRu = d.date.slice(8, 10) + '.' + d.date.slice(5, 7) + '.' + d.date.slice(0, 4);
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
    rec[O['Создан']] = txt_(nowStr_());
    rec[O['Статус']] = ST.new;
    rec[O['Букет']] = txt_(item.n);
    rec[O['Размер']] = d.size;
    rec[O['Букет, ₽']] = bouquet;
    rec[O['Доставка, ₽']] = delivery;
    rec[O['Итого, ₽']] = total;
    rec[O['Дата']] = txt_(dateRu);
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
    rec[O['Обновлён']] = txt_(nowStr_());
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

  return { ok: true, no: no, total: total, key: key, bot: botName_() };
}

function upsertClient_(phone, name, sum, last, tgId) {
  const sh = sheet_(SH.cli);
  const v = sh.getDataRange().getValues();
  const digits = phone.replace(/\D/g, '');
  const today = Utilities.formatDate(new Date(), TZ, 'dd.MM.yyyy');
  for (let i = 1; i < v.length; i++) {
    if (String(v[i][0]).replace(/\D/g, '') === digits) {
      if (sum) {
        sh.getRange(i + 1, 2, 1, 6).setValues([[txt_(name), (Number(v[i][2]) || 0) + 1,
          (Number(v[i][3]) || 0) + sum, txt_(String(v[i][4])), txt_(today), txt_(last)]]);
      }
      if (tgId) sh.getRange(i + 1, 8).setValue(txt_(tgId));
      return;
    }
  }
  if (sum) sh.appendRow([txt_(phone), txt_(name), 1, sum, txt_(today), txt_(today), txt_(last), txt_(tgId || '')]);
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
  say_(chat, 'Здравствуйте! Это бот ' + shop + ' 🌸\nЗдесь мы присылаем статусы заказов.\n\nВыбрать букет: ' + SITE_URL);
}

function linkBuyer_(chat, from, no, key) {
  const row = findOrder_(no);
  if (!row) return say_(chat, 'Не нашли такой заказ. Позвоните нам, пожалуйста.');
  const sh = sheet_(SH.ord);
  const r = sh.getRange(row, 1, 1, ORD_H.length).getValues()[0];
  if (String(r[O['тех: ключ']]) !== key) return say_(chat, 'Ссылка устарела. Позвоните нам, пожалуйста.');
  const handle = (from && from.username) ? '@' + from.username : String(chat);
  sh.getRange(row, O['Telegram покупателя'] + 1).setValue(txt_(chat + ' ' + handle));
  upsertClient_(String(r[O['Телефон']]), String(r[O['Покупатель']]), 0, '', handle);
  const o = orderAt_(row);
  say_(chat, 'Здравствуйте, ' + esc_(o.name.split(' ')[0]) + '! 🌸\n\n' +
    'Ваш заказ <b>№' + o.no + '</b>\n' + o.item + ', размер ' + o.size + '\n' +
    whenText_(o) + '\nИтого: ' + money_(o.total) + '\n\n' +
    'Статус: ' + ST_ICON[o.code] + ' ' + o.status + '\n\nСюда придёт сообщение, когда статус изменится.');
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
    sh.getRange(row, O['Обновлён'] + 1).setValue(txt_(nowStr_()));
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
      : 'Ваш букет собран 💐' + (o.photo ? ' Фото пришлём перед отправкой.' : ' Скоро передадим курьеру.'),
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
    return String(x.r[O['Дата']]) === ru && String(x.r[O['Статус']]) !== ST.cancel;
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
  const flags = [];
  if (o.photo) flags.push('📷 фото до отправки');
  if (o.anon) flags.push('🕶 анонимно');
  if (flags.length) L.push(flags.join(' · '));
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
    date: String(r[O['Дата']]), interval: String(r[O['Интервал']]),
    slot: Number(String(r[O['Интервал']]).split(':')[0]) || 0,
    pickup: String(r[O['Получение']]) === 'Самовывоз', address: String(r[O['Адрес']]),
    name: String(r[O['Покупатель']]), phone: String(r[O['Телефон']]),
    recipient: String(r[O['Получатель']]), rphone: String(r[O['Телефон получателя']]),
    card: String(r[O['Открытка']]), photo: String(r[O['Фото до отправки']]) === 'Да', anon: String(r[O['Анонимно']]) === 'Да',
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
