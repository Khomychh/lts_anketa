/**
 * Львівська Богословська Семінарія — приймання вступних анкет.
 *
 * Приймає POST від index.html і дописує рядок до аркуша «Заявки».
 * Аркуш і рядок заголовків створюються автоматично при першому запуску.
 *
 * Публікація: Deploy → New deployment → Web app
 *   Execute as:      Me
 *   Who has access:  Anyone
 * Отриманий URL (…/exec) вставте в index.html у константу GOOGLE_SCRIPT_URL.
 */

var SHEET_NAME = 'Заявки';

/** Порядок колонок таблиці. key — атрибут name у формі, header — заголовок колонки. */
var FIELDS = [
  { key: 'full_name',           header: 'ПІБ' },
  { key: 'birth_date',          header: 'Дата народження' },
  { key: 'phone',               header: 'Телефон' },
  { key: 'email',               header: 'Email' },
  { key: 'city',                header: 'Місто/область' },
  { key: 'photo',               header: 'Фото' },
  { key: 'school',              header: 'Заклад освіти' },
  { key: 'graduation_year',     header: 'Рік закінчення / курс' },
  { key: 'specialty',           header: 'Спеціальність' },
  { key: 'gpa',                 header: 'Середній бал' },
  { key: 'program',             header: 'Програма' },
  { key: 'study_form',          header: 'Форма навчання' },
  { key: 'level',               header: 'Рівень' },
  { key: 'source',              header: 'Звідки дізналися' },
  { key: 'applied_before',      header: 'Подавали документи раніше' },
  { key: 'motivation',          header: 'Чому саме тут' },
  { key: 'application',         header: 'Застосування освіти' },
  { key: 'ministry_experience', header: 'Досвід служіння' },
  { key: 'pastor_name',         header: 'ПІБ наставника' },
  { key: 'pastor_contact',      header: 'Контакт наставника' },
  { key: 'church',              header: 'Церква/громада' },
  { key: 'dormitory',           header: 'Потребує гуртожиток' },
  { key: 'financial_aid',       header: 'Потрібна фінансова допомога' },
  { key: 'contact_method',      header: 'Спосіб зв’язку' },
  { key: 'consent',             header: 'Згода на обробку даних' },
  { key: 'gpa_scale',           header: 'Шкала оцінювання' }
];

/** Папка Google Drive для фото. Створюється автоматично. */
var PHOTO_FOLDER = 'Анкети — фото';

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var p = (e && e.parameter) || {};

    // Спрацював honeypot — тихо ігноруємо запит.
    if (String(p.website || '').trim() !== '') {
      return json({ result: 'success' });
    }

    var sheet = getSheet();
    var row = [new Date()];

    for (var i = 0; i < FIELDS.length; i++) {
      var key = FIELDS[i].key;
      var value = p[key] == null ? '' : String(p[key]);
      if (key === 'photo') value = savePhoto(value, p.full_name);
      row.push(value);
    }

    sheet.appendRow(row);
    return json({ result: 'success' });
  } catch (err) {
    return json({ result: 'error', message: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/** Перевірка, що веб-застосунок опубліковано: відкрийте URL у браузері. */
function doGet() {
  return json({ result: 'success', message: 'Приймання анкет працює.' });
}

function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

  var headers = ['Час надходження'];
  for (var i = 0; i < FIELDS.length; i++) headers.push(FIELDS[i].header);

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  } else if (sheet.getLastColumn() < headers.length) {
    // у формі зʼявилися нові поля — дописуємо заголовки в кінець,
    // наявні колонки й рядки при цьому не зсуваються
    var from = sheet.getLastColumn();
    sheet.getRange(1, from + 1, 1, headers.length - from)
         .setValues([headers.slice(from)]);
  }
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  return sheet;
}

/**
 * Фото приходить як data URL (стиснене на клієнті). Кладемо файл на Drive
 * і повертаємо посилання; якщо фото немає — порожній рядок.
 */
function savePhoto(dataUrl, name) {
  if (!dataUrl || dataUrl.indexOf('data:image/') !== 0) return '';
  try {
    var parts = dataUrl.split(',');
    var mime = parts[0].slice(5).split(';')[0];
    var bytes = Utilities.base64Decode(parts[1]);
    var safeName = String(name || 'anketa').replace(/[^\p{L}\p{N} _-]/gu, '').trim() || 'anketa';
    var blob = Utilities.newBlob(bytes, mime, safeName + '.jpg');

    var folders = DriveApp.getFoldersByName(PHOTO_FOLDER);
    var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(PHOTO_FOLDER);
    return folder.createFile(blob).getUrl();
  } catch (err) {
    return 'помилка збереження фото: ' + err;
  }
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
