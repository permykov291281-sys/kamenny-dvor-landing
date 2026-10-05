/**
 * Приём заявок с лендинга «Каменный Двор» (укладка брусчатки) в Google-таблицу.
 *
 * ID таблицы руками искать не нужно: скрипт сам находит её по имени,
 * а если не нашёл — создаёт и запоминает. Ссылку на таблицу всегда можно
 * получить, открыв /exec в браузере.
 *
 * КАК РАЗВЕРНУТЬ (5 минут):
 *  1. script.google.com → «Новый проект», назвать «Заявки — брусчатка».
 *  2. Выделить весь код-заглушку, вставить этот файл, сохранить (Ctrl+S).
 *  3. Развернуть → Новое развёртывание → тип «Веб-приложение»:
 *       Запуск от имени: «Я»
 *       Доступ:          «У всех»  ← обязательно, иначе форма получит 401
 *     → Развернуть.
 *  4. Google попросит разрешения (Таблицы, Диск, Почта) —
 *     «Дополнительно» → «Перейти к проекту» → «Разрешить».
 *  5. Скопировать ссылку вида https://script.google.com/macros/s/AKfy.../exec
 *     и прислать её мне — я подставлю её в index.html в FORM_ENDPOINT.
 *  6. Проверка: открыть эту же ссылку в браузере. Должно быть
 *     {"ok":true, ...} и прямая ссылка на таблицу с заявками.
 *
 * ЕСЛИ ПОТОМ МЕНЯТЬ КОД: Развернуть → Управление развёртываниями → карандаш →
 * Версия: «Создать» → Развернуть. Тогда ссылка /exec останется прежней.
 */

const SHEET_TITLE   = 'сайт брусчатка';                       // как называется таблица
const SHEET_NAME    = 'Заявки';                               // лист внутри неё
const NOTIFY_EMAIL  = 'ваша-почта@example.com';    // дубль заявки на почту, '' — выключить

const HEADERS = ['Дата и время', 'Имя', 'Телефон', 'Объект', 'Сроки', 'Расчёт с сайта', 'Комментарий', 'Источник'];

/* ---------- приём заявки с сайта ---------- */

function doPost(e) {
  try {
    const data  = JSON.parse(e.postData.contents);
    const sheet = getSheet_();

    sheet.appendRow([
      new Date(),
      String(data.name    || '').slice(0, 120),
      String(data.phone   || '').slice(0, 40),
      String(data.object  || '').slice(0, 80),
      String(data.when    || '').slice(0, 60),
      String(data.calc    || '').slice(0, 200),
      String(data.comment || '').slice(0, 600),
      String(data.source  || 'сайт').slice(0, 60)
    ]);

    notify_(data);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

/* ---------- открыть /exec в браузере: проверка + ссылка на таблицу ---------- */

function doGet() {
  try {
    const sheet = getSheet_();
    return json_({
      ok: true,
      service: 'kamenny-dvor',
      spreadsheet: sheet.getParent().getName(),
      url: sheet.getParent().getUrl(),
      rows: Math.max(0, sheet.getLastRow() - 1),
      notify: NOTIFY_EMAIL || 'выключено'
    });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

/* ---------- письмо-дубль ---------- */

function notify_(data) {
  if (!NOTIFY_EMAIL) return;
  try {
    MailApp.sendEmail({
      to: NOTIFY_EMAIL,
      subject: 'Заявка с сайта «Каменный Двор»: ' + (data.name || 'без имени'),
      htmlBody:
        '<div style="font-family:Arial,sans-serif;font-size:15px;line-height:1.6;color:#23262B">' +
        '<h2 style="margin:0 0 12px">Новая заявка на замер</h2>' +
        row_('Имя',            data.name)    +
        row_('Телефон',        data.phone)   +
        row_('Что мостим',     data.object)  +
        row_('Сроки',          data.when)    +
        row_('Расчёт с сайта', data.calc)    +
        row_('Комментарий',    data.comment) +
        row_('Откуда нажали',  data.source)  +
        '</div>'
    });
  } catch (err) {
    // письмо — не главное; заявка уже в таблице, молча продолжаем
    console.warn('Письмо не ушло: ' + err);
  }
}

function row_(label, value) {
  return '<p style="margin:4px 0"><b>' + label + ':</b> ' + (value || '—') + '</p>';
}

/* ---------- поиск таблицы: свойства → активная → по имени на Диске → создать ---------- */

function getSheet_() {
  const props = PropertiesService.getScriptProperties();
  let ss = null;

  const savedId = props.getProperty('SPREADSHEET_ID');
  if (savedId) {
    try { ss = SpreadsheetApp.openById(savedId); } catch (err) { ss = null; }
  }

  if (!ss) {
    // если скрипт создан из самой таблицы — она и есть активная
    try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (err) { ss = null; }
  }

  if (!ss) {
    // ищем на Диске по имени
    const found = DriveApp.getFilesByName(SHEET_TITLE);
    if (found.hasNext()) ss = SpreadsheetApp.open(found.next());
  }

  if (!ss) {
    // не нашли — заводим свою
    ss = SpreadsheetApp.create(SHEET_TITLE);
  }

  props.setProperty('SPREADSHEET_ID', ss.getId());

  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 160);
    sheet.setColumnWidth(6, 260);
    sheet.setColumnWidth(7, 320);
  }
  return sheet;
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
