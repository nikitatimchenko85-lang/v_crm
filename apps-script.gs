/**
 * VILLATIC CRM · связка с Google Таблицей
 *
 * Установка:
 *  1. Создайте таблицу → Расширения → Apps Script.
 *  2. Вставьте этот код, сохраните.
 *  3. Развернуть → Новое развёртывание → тип «Веб-приложение».
 *     Запуск от имени: я.  Доступ: у кого есть ссылка.
 *  4. Скопируйте ссылку вида .../exec и вставьте её в CRM → Настройки.
 *  5. Один раз запустите функцию setupDailyBackup() — она включит
 *     ежедневную копию таблицы в папку Google Диска «Виллатик / CRM / CRM бэкапы».
 *
 * Лист «Сделки» — по одной строке на событие, читается глазами.
 * Лист «Платежи» — плоская расшифровка всех поступлений для отчётов и сводных таблиц.
 * Лист «Архив» — 30 последних версий листа «Сделки», на случай кривой выгрузки.
 */

var SHEET = 'Сделки';
var PAYSHEET = 'Платежи';

var COLS = [
  ['id','id'], ['created','Создано'], ['statusId','Статус id'], ['status','Статус'],
  ['client','Клиент'], ['phone','Телефон'], ['contact','Контакт'], ['source','Источник'],
  ['type','Тип события'], ['venueId','Площадка id'], ['venue','Площадка'], ['whole','Вся территория'], ['plot','Площадка регистрации'], ['plotAmount','Сумма площадки'], ['holdUntil','Срок раздумья'], ['cancelDate','Отменено'], ['retained','Удержано'],
  ['refunded','Возвращено'], ['cancelReason','Причина отмены'],
  ['date','Дата'], ['time','Время'], ['timeStart','Начало'], ['timeEnd','Конец'],
  ['guests','Гостей'], ['amount','Сумма'], ['paid','Оплачено'], ['due','Остаток'],
  ['contractNo','Договор №'], ['contractDate','Дата договора'],
  ['nextDate','След. шаг дата'], ['nextText','След. шаг'], ['notes','Заметки'],
  ['payments','payments_json'], ['history','history_json'], ['party','party_json'], ['flags','flags_json'], ['plan','plan_json']
];

function doPost(e) {
  // webhook от АТС или формы — без блокировки, это быстрые записи
  if (e.parameter && e.parameter.hook) return handleHook(e.parameter.hook, e);

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var body = JSON.parse(e.postData.contents);
    if (body.action === 'notify') return out({status:'ok', sent: tgSendAll(body.messages || [])});
    if (body.action !== 'sync') return out({status:'error', message:'неизвестное действие'});

    // защита от обвала: в таблице было много записей, а пришло сильно меньше
    var sheetNow = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET);
    var had = sheetNow ? Math.max(0, sheetNow.getLastRow() - 1) : 0;
    var got = (body.deals || []).length;
    if (!body.force && had >= 5 && got < had * 0.6) {
      return out({status:'shrink', message:'В таблице ' + had + ' заявок, а это устройство прислало ' + got + '.'});
    }

    // защита от затирания: если в таблице данные новее, просим подтверждение
    var props = PropertiesService.getScriptProperties();
    var lastStamp = Number(props.getProperty('lastStamp') || 0);
    var stamp = Number(body.stamp || 0);
    if (!body.force && lastStamp && stamp && stamp < lastStamp) {
      return out({status:'stale', message:'В таблице есть более свежие данные от ' +
        Utilities.formatDate(new Date(lastStamp), Session.getScriptTimeZone(), 'dd.MM.yyyy HH:mm') + '.'});
    }

    var sh = sheet(SHEET);

    // защита от затирания: пришло сильно меньше строк, чем было
    var had = Math.max(0, sh.getLastRow() - 1);
    var now = body.deals.length;
    if (!body.force && had >= 4 && now < had / 2) {
      return out({status:'confirm',
        message:'В таблице ' + had + ' записей, а устройство прислало ' + now + '.'});
    }
    if (had) snapshot(sh, had);
    sh.clear();
    sh.getRange(1, 1, 1, COLS.length).setValues([COLS.map(function(c){ return c[1]; })])
      .setFontWeight('bold').setBackground('#ECEEEB');
    sh.setFrozenRows(1);

    var rows = body.deals.map(function(d) {
      return COLS.map(function(c){ return d[c[0]] === undefined ? '' : d[c[0]]; });
    });
    if (rows.length) sh.getRange(2, 1, rows.length, COLS.length).setValues(rows);

    // деньги — числовым форматом, даты — текстом ISO, чтобы не поехали
    MONEY_FMT = '#,##0 "' + (body.currency || '\u20BD') + '"';
    var money = MONEY_FMT;
    var iAmount = idx('amount') + 1;
    if (rows.length) sh.getRange(2, iAmount, rows.length, 3).setNumberFormat(money);
    hide(sh, ['id','statusId','venueId','payments','history','party','flags','plan']);
    sh.autoResizeColumns(idx('client') + 1, 4);

    writePayments(body.deals);
    writeLog(body.log || []);
    writeShifts(body.shifts || []);
    writeTasks(body.tasks || []);
    writeRentals(body.rentals || []);
    writeOps(body.ops || [], body.catNames || {}, body.accNames || {});
    var keepSettings = acceptSettings(body);
    writeBlobs({
      settings: keepSettings.settings,
      deleted: body.deleted || '',
      tasks: JSON.stringify(body.tasks || []),
      rentals: JSON.stringify(body.rentals || []),
      ops: JSON.stringify(body.ops || []),
      log: JSON.stringify(body.log || []),
      shifts: JSON.stringify(body.shifts || [])
    });
    props.setProperty('settingsStamp', String(keepSettings.stamp));
    markLeadsTaken(body.leadsSeen || []);
    var sentCount = tgSendAll(body.notify || []);
    // календари: ошибка здесь не должна ломать синхронизацию
    try { syncCalendars(body.deals || [], body.deleted); } catch (calErr) { Logger.log('Календари: ' + calErr); }
    try { writeBusySheet(body.deals || []); } catch (busyErr) { Logger.log('Занятость: ' + busyErr); }
    rememberSent((body.notify || []).map(function(n) { return n.key; }));
    props.setProperty('lastStamp', String(stamp || Date.now()));
    return out({status:'ok', rows: rows.length, sent: sentCount});
  } catch (err) {
    return out({status:'error', message: String(err)});
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

/**
 * Снимок предыдущего состояния перед перезаписью.
 * Хранит 30 последних версий на листе «Архив» — страховка от кривой выгрузки.
 */
function snapshot(sh, rows) {
  try {
    var arc = sheet('Архив');
    if (arc.getLastRow() === 0) {
      arc.getRange(1, 1, 1, 3).setValues([['Время','Записей','Данные (JSON)']])
         .setFontWeight('bold').setBackground('#ECEEEB');
      arc.setFrozenRows(1);
      arc.setColumnWidth(3, 420);
    }
    var data = sh.getRange(1, 1, rows + 1, sh.getLastColumn()).getValues();
    arc.insertRowAfter(1);
    arc.getRange(2, 1, 1, 3).setValues([[
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM.yyyy HH:mm'),
      rows,
      JSON.stringify(data).slice(0, 45000)
    ]]);
    var extra = arc.getLastRow() - 31;
    if (extra > 0) arc.deleteRows(32, extra);
  } catch (err) {}
}

/**
 * Ежедневная копия всей таблицы в папку Google Диска — только если таблица
 * менялась с прошлой копии. Хранятся 14 последних копий.
 * Запустите setupDailyBackup() один раз вручную — дальше по расписанию.
 */
var ARCHIVE_SHEET = '⚠ АРХИВ';

/* true, если скрипт запущен в резервной копии, а не в рабочей таблице */
function isArchive() {
  return !!SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ARCHIVE_SHEET);
}
function refuseInArchive() {
  if (isArchive()) throw new Error('Это резервная копия. Триггеры запускаются только в рабочей таблице.');
}

function setupDailyBackup() {
  refuseInArchive();
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'dailyBackup') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('dailyBackup').timeBased().atHour(4).everyDays(1).create();
  dailyBackup();
}

function dailyBackup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var src = DriveApp.getFileById(ss.getId());

  // копируем, только если таблица менялась с прошлой копии:
  // в межсезонье папка перестаёт расти одинаковыми файлами
  var props = PropertiesService.getScriptProperties();
  var changedAt = src.getLastUpdated().getTime();
  var copiedAt = Number(props.getProperty('BACKUP_SRC_STAMP') || 0);
  if (changedAt <= copiedAt) return 'Изменений не было, копия не нужна';

  var folder = backupFolder();
  var day = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var name = 'АРХИВ CRM ' + day;
  var same = folder.getFilesByName(name);
  while (same.hasNext()) same.next().setTrashed(true);
  var copy = src.makeCopy(name, folder);
  props.setProperty('BACKUP_SRC_STAMP', String(changedAt));
  markArchive(copy.getId(), day);

  // держим 14 последних копий
  var files = [], it = folder.getFiles();
  while (it.hasNext()) files.push(it.next());
  files.sort(function(a, b){ return b.getDateCreated() - a.getDateCreated(); });
  files.slice(14).forEach(function(f){ f.setTrashed(true); });
}

/* Папка копий: Виллатик / CRM / CRM бэкапы.
 * Ищем по идентификатору, а не по имени, — чтобы не спутать с одноимённой папкой в другом месте. */
var BACKUP_FOLDER_ID = '1dmPh6i9PCTvjMUCGhIJfI5F7HaZE6Ijv';   // CRM бэкапы
var BACKUP_PARENT_ID = '1nPbJNHeDLdNzRchgIM41WQ__6RahKE_P';   // Виллатик / CRM

function backupFolder() {
  try {
    var f = DriveApp.getFolderById(BACKUP_FOLDER_ID);
    if (!f.isTrashed()) return f;
  } catch (e) {}
  // папку удалили — создаём заново на прежнем месте
  try { return DriveApp.getFolderById(BACKUP_PARENT_ID).createFolder('CRM бэкапы'); } catch (e) {}
  return DriveApp.createFolder('CRM бэкапы');
}

/* первым листом копии — крупное предупреждение */
function markArchive(id, day) {
  try {
    var cs = SpreadsheetApp.openById(id);
    var warn = cs.insertSheet(ARCHIVE_SHEET, 0);
    warn.getRange('A1').setValue('РЕЗЕРВНАЯ КОПИЯ — НЕ РЕДАКТИРОВАТЬ')
      .setFontSize(22).setFontWeight('bold').setFontColor('#B2451E');
    warn.getRange('A3').setValue('Копия рабочей таблицы CRM на ' + day.split('-').reverse().join('.') + '. ' +
      'Приложение сюда не пишет: правки в этой копии никуда не попадут.');
    warn.getRange('A4').setValue('Работайте в основной таблице. Копия нужна только для восстановления, ' +
      'если основная таблица пропала или испорчена.');
    warn.getRange('A5').setValue('Развёртывание и триггеры в копии не создавайте.');
    warn.getRange('A3:A5').setFontSize(12).setWrap(true);
    warn.setColumnWidth(1, 820);
    warn.setTabColor('#B2451E');
    cs.setActiveSheet(warn);
  } catch (err) {}
}

function doGet(e) {
  try {
    if (e.parameter && e.parameter.hook) return handleHook(e.parameter.hook, e);
    var act = e.parameter.action;
    if (act !== 'list' && act !== 'full') return out({status:'error', message:'неизвестное действие'});
    var sh = sheet(SHEET);
    var values = sh.getDataRange().getValues();
    if (values.length < 2) return out({status:'ok', deals: []});
    var deals = [];
    for (var r = 1; r < values.length; r++) {
      if (!values[r][0]) continue;
      var obj = {};
      COLS.forEach(function(c, i) {
        var v = values[r][i];
        if (c[0] === 'date' || c[0] === 'nextDate') v = asDate(v);
        obj[c[0]] = v;
      });
      deals.push(obj);
    }
    if (act === 'list') return out({status:'ok', deals: deals});

    var blobs = readBlobs();
    return out({status:'ok', deals: deals,
      settings: blobs.settings || '',
      settingsStamp: Number(PropertiesService.getScriptProperties().getProperty('settingsStamp') || 0),
      deleted: blobs.deleted || '[]',
      tasks: blobs.tasks || '[]',
      rentals: blobs.rentals || '[]',
      ops: blobs.ops || '[]',
      log: blobs.log || '[]',
      calls: readCalls(),
      leads: readLeads(),
      shifts: blobs.shifts || '[]'});
  } catch (err) {
    return out({status:'error', message: String(err)});
  }
}

var MONEY_FMT = '#,##0 "\u20BD"';

function writePayments(deals) {
  var sh = sheet(PAYSHEET);
  sh.clear();
  sh.getRange(1, 1, 1, 6).setValues([['Дата','Клиент','Событие','Площадка','Назначение','Сумма']])
    .setFontWeight('bold').setBackground('#ECEEEB');
  sh.setFrozenRows(1);
  var rows = [];
  deals.forEach(function(d) {
    var pays = [];
    try { pays = JSON.parse(d.payments || '[]'); } catch (err) {}
    pays.forEach(function(p) {
      rows.push([p.date || '', d.client || '', d.type || '', d.venue || '', p.kind || 'Платёж', Number(p.sum) || 0]);
    });
  });
  rows.sort(function(a, b){ return String(a[0]).localeCompare(String(b[0])); });
  if (rows.length) {
    sh.getRange(2, 1, rows.length, 6).setValues(rows);
    sh.getRange(2, 6, rows.length, 1).setNumberFormat(MONEY_FMT);
  }
}

function writeLog(entries) {
  var sh = sheet('Журнал');
  sh.clear();
  sh.getRange(1, 1, 1, 6).setValues([['Время','Сотрудник','Заявка','Действие','Было','Стало']])
    .setFontWeight('bold').setBackground('#ECEEEB');
  sh.setFrozenRows(1);
  if (!entries.length) return;
  var rows = entries.map(function(x) {
    return [Utilities.formatDate(new Date(x.t), Session.getScriptTimeZone(), 'dd.MM.yyyy HH:mm'),
            x.u || '', x.name || '', x.f ? 'изменено: ' + x.f : (x.a || ''), x.from || '', x.to || ''];
  }).reverse();
  sh.getRange(2, 1, rows.length, 6).setValues(rows);
}

function writeShifts(list) {
  var sh = sheet('Смены');
  sh.clear();
  sh.getRange(1, 1, 1, 7).setValues([['Дата','Сотрудник','Начало','Конец','Минут','Действий','Устройство']])
    .setFontWeight('bold').setBackground('#ECEEEB');
  sh.setFrozenRows(1);
  if (!list.length) return;
  var tz = Session.getScriptTimeZone();
  var rows = list.map(function(s) {
    var end = s.end || s.last;
    return [s.day || '', s.name || '',
            Utilities.formatDate(new Date(s.start), tz, 'HH:mm'),
            s.end ? Utilities.formatDate(new Date(end), tz, 'HH:mm') : 'открыта',
            Math.round((end - s.start) / 60000), s.actions || 0, s.dev || ''];
  }).reverse();
  sh.getRange(2, 1, rows.length, 7).setValues(rows);
}

function writeOps(list, catNames, accNames) {
  var sh = sheet('Операции');
  sh.clear();
  sh.getRange(1, 1, 1, 7).setValues([['Дата','Тип','Статья','Касса','Событие','Кому / за что','Сумма']])
    .setFontWeight('bold').setBackground('#ECEEEB');
  sh.setFrozenRows(1);
  if (!list.length) return;
  var kinds = {'in':'приход', 'out':'расход', 'move':'перевод'};
  var rows = list.map(function(o) {
    return [o.date || '', kinds[o.kind] || o.kind,
            o.kind === 'move' ? ('в ' + (accNames[o.toAccountId] || '')) : (catNames[o.catId] || ''),
            accNames[o.accountId] || '', o.dealTitle || '', o.party || '',
            (o.kind === 'out' ? -1 : 1) * (Number(o.sum) || 0)];
  }).sort(function(a, b) { return String(a[0]).localeCompare(String(b[0])); });
  sh.getRange(2, 1, rows.length, 7).setValues(rows);
  sh.getRange(2, 7, rows.length, 1).setNumberFormat(MONEY_FMT);
}

function writeRentals(list) {
  var sh = sheet('Прокат');
  sh.clear();
  sh.getRange(1, 1, 1, 9).setValues([['Выдача','Возврат','Суток','Кому','Тип','Событие','Позиции','Сумма','Статус']])
    .setFontWeight('bold').setBackground('#ECEEEB');
  sh.setFrozenRows(1);
  if (!list.length) return;
  var rows = list.map(function(o) {
    return [o.date || '', o.dateTo || o.date || '', Number(o.days) || 1, o.renterName || '',
            o.renterType === 'client' ? 'арендатор' : 'подрядчик',
            o.dealTitle || '', o.itemsText || '', Number(o.total) || 0, o.status || ''];
  }).sort(function(a, b) { return String(a[0]).localeCompare(String(b[0])); });
  sh.getRange(2, 1, rows.length, 9).setValues(rows);
  sh.getRange(2, 8, rows.length, 1).setNumberFormat(MONEY_FMT);
}

function writeTasks(list) {
  var sh = sheet('Задачи');
  sh.clear();
  sh.getRange(1, 1, 1, 7).setValues([['Срок','Время','Задача','Событие','Исполнитель','Статус','Выполнил']])
    .setFontWeight('bold').setBackground('#ECEEEB');
  sh.setFrozenRows(1);
  if (!list.length) return;
  var rows = list.map(function(t) {
    return [t.due || '', t.time || '', t.title || '', t.dealTitle || '', t.assigneeName || '',
            t.done ? 'выполнено' : 'в работе', t.doneBy || ''];
  }).sort(function(a, b) { return String(a[0] + a[1]).localeCompare(String(b[0] + b[1])); });
  sh.getRange(2, 1, rows.length, 7).setValues(rows);
}



/* ================= GOOGLE КАЛЕНДАРИ ================= */
/*
 * Два календаря:
 *   «Villatic · внутренний» — зал, статус, клиент, телефон, контакт, гости;
 *   «Villatic · занятость»  — только зал и статус, для общего доступа.
 * Запустите setupCalendars один раз: календари создадутся и заполнятся.
 * Дальше они обновляются сами при каждой выгрузке из CRM — только изменившиеся события.
 */
var CAL_MAP = 'Календари';
var CAL_INT_NAME = 'Villatic · внутренний';
var CAL_PUB_NAME = 'Villatic · занятость';

function setupCalendars() {
  refuseInArchive();
  var props = PropertiesService.getScriptProperties();
  var intCal = calendarNamed(CAL_INT_NAME, 'Заявки и брони VILLATIC: зал, статус, контакт клиента. Только для сотрудников.');
  props.setProperty('CAL_INT', intCal.getId());
  var n = syncCalendars(readDealObjs(), null, true);
  var msg = 'Внутренний календарь готов, событий: ' + n + '.';
  Logger.log(msg);
  return msg;
}

/* Запустите один раз: удаляет общий Google-календарь, его заменила страница занятости */
function removePublicCalendar() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('CAL_PUB');
  var cals = id ? [CalendarApp.getCalendarById(id)] : CalendarApp.getCalendarsByName(CAL_PUB_NAME);
  var n = 0;
  cals.forEach(function(c) { if (c) { c.deleteCalendar(); n++; } });
  props.deleteProperty('CAL_PUB');
  var msg = n ? 'Общий календарь удалён' : 'Общего календаря не было';
  Logger.log(msg);
  return msg;
}

/* ---------- лист «Занятость» для публичной страницы ----------
 * Только дата, зал и статус — без клиентов, телефонов и сумм.
 * Этот лист публикуется в интернет как CSV, остальные — нет. */
var BUSY_SHEET = 'Занятость';

function writeBusySheet(deals) {
  var from = dayShift(-1);
  var rows = [];
  deals.forEach(function(d) {
    var sid = String(d.statusId || '');
    if (sid === 'lost') return;
    var date = ymdOf(d.date);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(date) || date < from) return;
    var whole = d.whole === true || d.whole === 'да';
    rows.push([date, whole ? 'Вся территория' : String(d.venue || ''), publicStatus(sid)]);
  });
  rows.sort(function(a, b) { return a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : 0; });

  var sh = sheet(BUSY_SHEET);
  sh.clear();
  sh.getRange(1, 1, 1, 3).setValues([['date', 'hall', 'status']]);
  if (rows.length) {
    sh.getRange(2, 1, rows.length, 1).setNumberFormat('@');     // дата текстом, без перевода в формат таблицы
    sh.getRange(2, 1, rows.length, 3).setValues(rows);
  }
  return rows.length;
}

/* первое заполнение вручную */
function publishBusy() {
  refuseInArchive();
  var n = writeBusySheet(readDealObjs());
  var msg = 'Лист «Занятость» заполнен, записей: ' + n +
            '. Опубликуйте именно этот лист: Файл → Поделиться → Опубликовать в интернете → «Занятость» → CSV.';
  Logger.log(msg);
  return msg;
}

function calendarNamed(name, about) {
  var found = CalendarApp.getCalendarsByName(name);
  if (found.length) return found[0];
  return CalendarApp.createCalendar(name, {summary: about, timeZone: Session.getScriptTimeZone()});
}

/* статусы для общего календаря — без внутренней кухни */
function publicStatus(id) {
  if (id === 'book' || id === 'done') return 'Бронь';
  return 'Предварительно';
}
function statusColor(id) {
  if (id === 'book') return CalendarApp.EventColor.GREEN;
  if (id === 'done') return CalendarApp.EventColor.GRAY;
  return CalendarApp.EventColor.YELLOW;
}

function hhmm(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'HH:mm');
  var m = String(v || '').match(/(\d{1,2}):(\d{2})/);
  return m ? ('0' + m[1]).slice(-2) + ':' + m[2] : '';
}
function cleanTime(v) {
  if (v instanceof Date) return '';
  var m = String(v || '').match(/^(\d{1,2}):(\d{2})$/);
  return m ? ('0' + m[1]).slice(-2) + ':' + m[2] : '';
}
function ymdOf(v) { return v instanceof Date ? ymd(v) : String(v || '').slice(0, 10); }

/* что должно стоять в календаре по заявке; null — ничего */
function calPlan(d) {
  var statusId = String(d.statusId || '');
  var date = ymdOf(d.date);
  if (!date || !/^\d{4}-\d{2}-\d{2}$/.test(date) || statusId === 'lost') return null;
  var whole = d.whole === true || d.whole === 'да';
  var hall = whole ? 'Вся территория' : String(d.venue || 'Зал');
  var status = String(d.status || '');
  var client = String(d.client || '').trim();
  // время из таблицы ненадёжно (сдвиг дат 1899 года), берём только строки вида 16:00
  var ts = cleanTime(d.timeStart), te = cleanTime(d.timeEnd);
  var lines = [];
  if (client) lines.push('Клиент: ' + client);
  if (d.phone) lines.push('Телефон: ' + d.phone);
  if (d.contact) lines.push('Контакт: ' + d.contact);
  if (ts) lines.push('Время: ' + ts + (te ? '–' + te : ''));
  if (d.guests) lines.push('Гостей: ' + d.guests);
  if (d.plot === true || d.plot === 'да') lines.push('С поляной для выездной регистрации');
  lines.push('Статус: ' + status);
  return {
    date: date, start: '', end: '', statusId: statusId,     // всегда на весь день
    intTitle: hall + ' · ' + status + (client ? ' · ' + client : ''),
    intText: lines.join('\n'),
    pubTitle: hall + ' · ' + publicStatus(statusId)
  };
}

function placeEvent(cal, id, title, text, p) {
  var ev = null;
  if (id) { try { ev = cal.getEventById(id); } catch (e) { ev = null; } }
  var tz = Session.getScriptTimeZone();
  var allDay = !p.start;
  var start, end;
  if (!allDay) {
    start = Utilities.parseDate(p.date + ' ' + p.start, tz, 'yyyy-MM-dd HH:mm');
    end = p.end ? Utilities.parseDate(p.date + ' ' + p.end, tz, 'yyyy-MM-dd HH:mm')
                : new Date(start.getTime() + 3 * 3600 * 1000);
    if (end <= start) end = new Date(end.getTime() + 24 * 3600 * 1000);   // до 00:00 — это следующий день
  }
  var day = Utilities.parseDate(p.date, tz, 'yyyy-MM-dd');

  // сменился тип (весь день ↔ по времени) — проще пересоздать
  if (ev && ev.isAllDayEvent() !== allDay) { try { ev.deleteEvent(); } catch (e) {} ev = null; }
  if (!ev) {
    ev = allDay ? cal.createAllDayEvent(title, day, {description: text})
                : cal.createEvent(title, start, end, {description: text});
  } else {
    ev.setTitle(title);
    ev.setDescription(text);
    if (allDay) ev.setAllDayDate(day); else ev.setTime(start, end);
  }
  try { ev.setColor(statusColor(p.statusId)); } catch (e) {}
  return ev.getId();
}

function dropEvent(cal, id) {
  if (!id) return;
  try { var ev = cal.getEventById(id); if (ev) ev.deleteEvent(); } catch (e) {}
}

function syncCalendars(deals, deletedJson, full) {
  var props = PropertiesService.getScriptProperties();
  var intId = props.getProperty('CAL_INT'), pubId = props.getProperty('CAL_PUB');
  if (!intId) return 0;                                   // календарь ещё не подключён
  var intCal = CalendarApp.getCalendarById(intId);
  var pubCal = pubId ? CalendarApp.getCalendarById(pubId) : null;   // общий больше не ведём
  if (!intCal) return 0;

  // карта: заявка → события и отпечаток содержимого
  var sh = sheet(CAL_MAP), map = {};
  if (sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues().forEach(function(r) {
      if (r[0]) map[String(r[0])] = {int: String(r[1] || ''), pub: String(r[2] || ''), hash: String(r[3] || '')};
    });
  }

  var dead = {};
  try { (JSON.parse(deletedJson || '[]') || []).forEach(function(t){ dead[t.id || t] = 1; }); } catch (e) {}

  var seen = {}, touched = 0;
  deals.forEach(function(d) {
    var id = String(d.id || '');
    if (!id) return;
    seen[id] = 1;
    var plan = dead[id] ? null : calPlan(d);
    var cur = map[id];
    if (!plan) {
      if (cur) { dropEvent(intCal, cur.int); if (pubCal) dropEvent(pubCal, cur.pub); delete map[id]; touched++; }
      return;
    }
    var hash = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5,
                 JSON.stringify(plan), Utilities.Charset.UTF_8));
    if (cur && cur.hash === hash && !full) return;           // ничего не менялось
    map[id] = {
      int: placeEvent(intCal, cur && cur.int, plan.intTitle, plan.intText, plan),
      pub: pubCal ? placeEvent(pubCal, cur && cur.pub, plan.pubTitle, '', plan) : '',
      hash: hash
    };
    touched++;
  });

  // заявки, которых больше нет в CRM, убираем из календарей
  Object.keys(map).forEach(function(id) {
    if (!seen[id]) { dropEvent(intCal, map[id].int); if (pubCal) dropEvent(pubCal, map[id].pub); delete map[id]; touched++; }
  });

  sh.clear();
  var rows = [['Заявка', 'Событие внутр.', 'Событие общ.', 'Отпечаток']];
  Object.keys(map).forEach(function(id) { rows.push([id, map[id].int, map[id].pub, map[id].hash]); });
  sh.getRange(1, 1, rows.length, 4).setValues(rows);
  if (!sh.isSheetHidden()) sh.hideSheet();
  return full ? Object.keys(map).length : touched;
}

/* ================= РАСПИСАНИЕ: УТРЕННЯЯ СВОДКА И ПРОВЕРКИ ================= */
/*
 * Запустите один раз функцию installTriggers — она создаст два триггера:
 *   • dailyDigest  — каждое утро в 9:00 присылает сводку дня
 *   • hourlyCheck  — раз в час проверяет истёкшие сроки раздумья
 * Часовой пояс берётся из настроек проекта Apps Script.
 */

function installTriggers() {
  refuseInArchive();
  removeTriggers();
  ScriptApp.newTrigger('dailyDigest').timeBased().atHour(9).everyDays(1).create();
  ScriptApp.newTrigger('hourlyCheck').timeBased().everyHours(1).create();
  return 'Триггеры установлены';
}

function removeTriggers() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    var f = t.getHandlerFunction();
    if (f === 'dailyDigest' || f === 'hourlyCheck') ScriptApp.deleteTrigger(t);
  });
  return 'Триггеры удалены';
}

/* ---------- защита от повторов ---------- */
function sentKeys() {
  try { return JSON.parse(PropertiesService.getScriptProperties().getProperty('SENT_KEYS') || '[]'); }
  catch (e) { return []; }
}
function rememberSent(keys) {
  if (!keys || !keys.length) return;
  var all = sentKeys().concat(keys);
  if (all.length > 400) all = all.slice(all.length - 400);
  PropertiesService.getScriptProperties().setProperty('SENT_KEYS', JSON.stringify(all));
}

/* ---------- чтение данных из листов ---------- */
function ymd(d) { return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
function dayShift(n) { return ymd(new Date(Date.now() + n * 86400000)); }
function asStr(v) { return v instanceof Date ? ymd(v) : String(v || ''); }
function ru(v) {                                  // 2026-09-06 -> 06.09
  var s = asStr(v);
  return /^\d{4}-\d{2}-\d{2}/.test(s) ? s.slice(8,10) + '.' + s.slice(5,7) : s;
}

function readDealObjs() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET);
  if (!sh || sh.getLastRow() < 2) return [];
  var v = sh.getDataRange().getValues(), res = [];
  for (var r = 1; r < v.length; r++) {
    if (!v[r][0]) continue;
    var o = {};
    COLS.forEach(function(c, i) { o[c[0]] = v[r][i]; });
    res.push(o);
  }
  return res;
}
function blobList(key) {
  try { return JSON.parse(readBlobs()[key] || '[]'); } catch (e) { return []; }
}
function money(n) { return (Math.round(Number(n) || 0)).toLocaleString('ru-RU') + ' \u20BD'; }
function held(d) { return d.statusId === 'book' || d.statusId === 'done'; }

/* ---------- утренняя сводка ---------- */
function dailyDigest() {
  var today = ymd(new Date()), tom = dayShift(1), week = dayShift(7);
  var deals = readDealObjs();
  var lines = ['Сводка на ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM, EEEE')];

  var todayEv = deals.filter(function(d) { return asStr(d.date) === today && held(d); });
  var tomEv   = deals.filter(function(d) { return asStr(d.date) === tom && held(d); });

  if (todayEv.length) {
    lines.push('', 'Сегодня:');
    todayEv.forEach(function(d) {
      lines.push('• ' + d.client + ' — ' + d.venue + (d.time ? ', ' + d.time : '') +
                 (d.guests ? ', ' + d.guests + ' гостей' : ''));
    });
  } else {
    lines.push('', 'Сегодня событий нет.');
  }
  if (tomEv.length) {
    lines.push('', 'Завтра:');
    tomEv.forEach(function(d) { lines.push('• ' + d.client + ' — ' + d.venue); });
  }

  var tasks = blobList('tasks').filter(function(t) { return !t.done && t.due && t.due <= today; });
  if (tasks.length) {
    lines.push('', 'Задачи:');
    tasks.slice(0, 12).forEach(function(t) {
      lines.push('• ' + t.title + (t.due < today ? ' (просрочена ' + ru(t.due) + ')' : ''));
    });
  }

  var expired = deals.filter(function(d) {
    return d.holdUntil && !held(d) && new Date(d.holdUntil) < new Date();
  });
  if (expired.length) {
    lines.push('', 'Истёк срок раздумья:');
    expired.forEach(function(d) { lines.push('• ' + d.client + ' — ' + asStr(d.date) + ', ' + d.venue); });
  }

  var debts = deals.filter(function(d) {
    var ds = asStr(d.date);
    return held(d) && ds >= today && ds <= week && Number(d.due) > 0;
  });
  if (debts.length) {
    lines.push('', 'Ждём оплату на неделе:');
    debts.forEach(function(d) { lines.push('• ' + d.client + ' — ' + money(d.due) + ' до ' + ru(d.date)); });
  }

  var rent = blobList('rentals').filter(function(o) {
    var to = o.dateTo && o.dateTo >= o.date ? o.dateTo : o.date;
    return o.date && o.date <= today && to >= today &&
           (o.status === 'reserved' || o.status === 'issued');
  });
  if (rent.length) {
    lines.push('', 'Прокат сегодня:');
    rent.forEach(function(o) {
      lines.push('• ' + (o.renterName || '') + ' — ' + (o.itemsText || '') +
                 (o.status === 'reserved' ? ' (выдать)' : ' (у клиента)'));
    });
  }

  tgSendAll([{text: lines.join('\n')}]);
  return lines.join('\n');
}

/* ---------- ежечасная проверка сроков ---------- */
function hourlyCheck() {
  var now = new Date(), already = sentKeys(), fresh = [], msgs = [];
  readDealObjs().forEach(function(d) {
    if (!d.holdUntil || held(d) || d.statusId === 'lost') return;
    var until = new Date(d.holdUntil);
    if (isNaN(until.getTime()) || until > now) return;
    var key = 'hold:' + d.id + ':' + d.holdUntil;
    if (already.indexOf(key) >= 0) return;
    fresh.push(key);
    msgs.push({text: 'Истёк срок раздумья: ' + d.client + ' — ' + ru(d.date) + ', ' + d.venue +
                     '. Позвоните клиенту или передайте очередь следующему.'});
  });
  if (msgs.length) { tgSendAll(msgs); rememberSent(fresh); }
  return msgs.length;
}

/* ================= ИНТЕГРАЦИИ: звонки, заявки, Telegram ================= */

var CALLS = 'Звонки';
var LEADS = 'Входящие';

function handleHook(kind, e) {
  // данные приходят либо формой, либо JSON — принимаем оба вида
  var p = {}, k;
  if (e.parameter) for (k in e.parameter) p[k] = e.parameter[k];
  if (e.postData && e.postData.contents) {
    try {
      var j = JSON.parse(e.postData.contents);
      for (k in j) p[k] = j[k];
    } catch (err) {}
  }

  // пишем в отладку ДО всех проверок — чтобы был виден любой пришедший запрос
  debugHook(kind, p);

  var need = PropertiesService.getScriptProperties().getProperty('HOOK_TOKEN');
  var got = (e.parameter && e.parameter.token) || '';
  if (need && got !== need) {
    debugHook('ОТКАЗ', {причина:'токен не совпал', пришёл:got, ожидался:need ? 'задан в свойствах' : 'не задан'});
    return out({status:'error', message:'неверный токен'});
  }

  if (kind === 'call') return saveCall(p);
  if (kind === 'lead') return saveLead(p);
  debugHook('ОТКАЗ', {причина:'неизвестный хук', hook:kind});
  return out({status:'error', message:'неизвестный хук'});
}

function pick(p, names) {
  for (var i = 0; i < names.length; i++) {
    var v = p[names[i]];
    if (v !== undefined && v !== null && String(v) !== '') return String(v);
  }
  return '';
}

function saveCall(p) {
  var sh = sheet(CALLS);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, 7).setValues([['Время','Направление','Номер','Секунд','Сотрудник','Запись','id']])
      .setFontWeight('bold').setBackground('#ECEEEB');
    sh.setFrozenRows(1);
  }
  // ключи под Sipuni, Zadarma, Mango, UIS и произвольные формы
  var dirRaw = pick(p, ['dir','direction','call_type','callType','type']);
  // у Сипуни type: 0 — входящий, 1 — исходящий
  var dir = (/out|исход/i.test(dirRaw) || dirRaw === '1') ? 'out' : 'in';
  var phone = pick(p, ['phone','src_num','caller_id','src','from','from_number','caller','client_number','number']);
  var staff = pick(p, ['staff','dst_num','answered','dst','to','called_did','internal','operator','user']);
  var dur   = pick(p, ['duration','dur','billsec','talk_time','seconds']);
  var rec   = pick(p, ['record','recording','rec','call_record','link','record_link']);
  var at    = pick(p, ['at','time','datetime','start','call_start','timestamp','date']);
  var id    = pick(p, ['call','id','call_id','callid','pbx_call_id','uid']) || Utilities.getUuid();

  var when = at ? new Date(at) : new Date();
  if (isNaN(when.getTime())) when = new Date();

  // события одного звонка (начало, ответ, завершение) сводим в одну строку
  var ids = sh.getLastRow() > 1
    ? sh.getRange(2, 7, sh.getLastRow() - 1, 1).getValues().map(function(r) { return String(r[0]); })
    : [];
  var at = ids.indexOf(String(id));

  if (at >= 0) {
    var row = at + 2;
    var cur = sh.getRange(row, 1, 1, 7).getValues()[0];
    var upd = [
      cur[0],                                             // время первого события
      dirRaw ? dir : cur[1],
      phone || cur[2],
      Math.max(Number(dur) || 0, Number(cur[3]) || 0),    // длительность приходит в последнем событии
      staff || cur[4],
      rec || cur[5],
      id
    ];
    sh.getRange(row, 1, 1, 7).setValues([upd]);
    return out({status:'ok', updated:true});
  }

  sh.appendRow([when, dir, phone, Number(dur) || 0, staff, rec, id]);
  return out({status:'ok'});
}

/* проверка без АТС: запустите testHook в редакторе — должна появиться строка в «Звонках» */
function testHook() {
  handleHook('call', {parameter:{
    token: PropertiesService.getScriptProperties().getProperty('HOOK_TOKEN') || '',
    call:'test-' + Date.now(), src_num:'79990001122', dst_num:'74950001122',
    type:'0', duration:'42', timestamp:Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss')
  }});
  return 'Готово: проверьте листы «Звонки» и «Отладка хуков»';
}

/* последние 100 сырых событий — чтобы видеть, какие поля шлёт АТС */
function debugHook(kind, p) {
  try {
    var sh = sheet('Отладка хуков');
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, 3).setValues([['Время','Хук','Что пришло']])
        .setFontWeight('bold').setBackground('#ECEEEB');
      sh.setFrozenRows(1);
    }
    sh.appendRow([new Date(), kind, JSON.stringify(p).slice(0, 5000)]);
    if (sh.getLastRow() > 101) sh.deleteRows(2, sh.getLastRow() - 101);
  } catch (err) {}
}

function readCalls() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CALLS);
  if (!sh || sh.getLastRow() < 2) return [];
  var from = Math.max(2, sh.getLastRow() - 499);
  var vals = sh.getRange(from, 1, sh.getLastRow() - from + 1, 7).getValues();
  return vals.map(function(r) {
    return {at: r[0] instanceof Date ? r[0].toISOString() : String(r[0]),
            dir: String(r[1]), phone: String(r[2]), dur: Number(r[3]) || 0,
            who: String(r[4] || ''), rec: String(r[5] || ''), id: String(r[6])};
  });
}

function saveLead(p) {
  var sh = sheet(LEADS);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, 10).setValues([['Время','Имя','Телефон','Контакт','Дата события','Гостей','Комментарий','Источник','id','Статус']])
      .setFontWeight('bold').setBackground('#ECEEEB');
    sh.setFrozenRows(1);
  }
  var id = pick(p, ['id','lead_id','uid']) || Utilities.getUuid();
  var name = pick(p, ['name','client','fio']);
  var phone = pick(p, ['phone','tel']);
  sh.appendRow([new Date(), name, phone,
    pick(p, ['contact','email','telegram']),
    pick(p, ['date','event_date']),
    pick(p, ['guests','pax']),
    pick(p, ['comment','message','text']),
    pick(p, ['source','utm_source']) || 'Сайт',
    id, 'новая']);
  tgSendAll([{text:'Новая заявка: ' + name + ' ' + phone}]);
  return out({status:'ok', id:id});
}

function readLeads() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LEADS);
  if (!sh || sh.getLastRow() < 2) return [];
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, 10).getValues();
  var res = [];
  vals.forEach(function(r) {
    if (String(r[9]) === 'принята') return;
    res.push({id:String(r[8]), name:String(r[1]), phone:String(r[2]), contact:String(r[3]),
      date: r[4] instanceof Date ? Utilities.formatDate(r[4], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(r[4] || ''),
      guests:String(r[5] || ''), comment:String(r[6] || ''), source:String(r[7] || 'Сайт')});
  });
  return res;
}

function markLeadsTaken(ids) {
  if (!ids || !ids.length) return;
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LEADS);
  if (!sh || sh.getLastRow() < 2) return;
  var range = sh.getRange(2, 9, sh.getLastRow() - 1, 2);
  var vals = range.getValues(), touched = false;
  vals.forEach(function(r) {
    if (ids.indexOf(String(r[0])) >= 0 && String(r[1]) !== 'принята') { r[1] = 'принята'; touched = true; }
  });
  if (touched) range.setValues(vals);
}

function tgSendAll(messages) {
  if (!messages || !messages.length) return 0;
  var props = PropertiesService.getScriptProperties();
  var token = props.getProperty('TG_TOKEN'), chats = props.getProperty('TG_CHAT');
  if (!token || !chats) return 0;
  var list = chats.split(',').map(function(x) { return x.trim(); }).filter(String);
  var sent = 0;
  messages.slice(0, 20).forEach(function(m) {
    list.forEach(function(chat) {
      try {
        UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/sendMessage', {
          method:'post', muteHttpExceptions:true,
          payload:{chat_id:chat, text:m.text, disable_web_page_preview:'true'}
        });
        sent++;
      } catch (err) {}
    });
  });
  return sent;
}

/* ---------- служебный лист: то, что нужно для одинаковых устройств ---------- */
var SVC = 'Служебное';
var CHUNK = 40000;

/* Решает, какие настройки оставить в таблице.
 * Пришедшие принимаются, только если они свежее сохранённых и не «беднее» их:
 * устройство с пустым списком сотрудников не может затереть настоящую команду. */
function staffCount(json) {
  try { return (JSON.parse(json || '{}').staff || []).length; } catch (e) { return 0; }
}
function acceptSettings(body) {
  var props = PropertiesService.getScriptProperties();
  var stored = readBlobs().settings || '';
  var storedStamp = Number(props.getProperty('settingsStamp') || 0);
  var incoming = body.settings || '';
  var incomingStamp = Number(body.settingsStamp || 0);

  if (!incoming) return {settings: stored, stamp: storedStamp};
  if (!stored || body.forceSettings) return {settings: incoming, stamp: Math.max(incomingStamp, Date.now())};

  var downgrade = staffCount(stored) > 1 && staffCount(incoming) <= 1;
  if (downgrade || incomingStamp <= storedStamp) return {settings: stored, stamp: storedStamp};
  return {settings: incoming, stamp: incomingStamp};
}

/* Восстановление настроек (команда, прайс, шаблоны) из резервной копии.
 * Запустите restoreSettings0912 один раз в редакторе — затем на всех устройствах
 * «Забрать настройки принудительно». Заявки при этом не трогаются. */
function restoreSettingsFromCopy(copyId) {
  refuseInArchive();
  var copy = SpreadsheetApp.openById(copyId).getSheetByName(SVC);
  if (!copy) throw new Error('В копии нет служебного листа');
  var vals = copy.getDataRange().getValues().filter(function(r){ return r[0] === 'settings'; });
  vals.sort(function(a, b) { return Number(a[1]) - Number(b[1]); });
  var settings = vals.map(function(r){ return String(r[2] || ''); }).join('');
  if (staffCount(settings) < 1) throw new Error('В копии не найдены настройки');

  var blobs = readBlobs();
  blobs.settings = settings;
  writeBlobs(blobs);
  PropertiesService.getScriptProperties().setProperty('settingsStamp', String(Date.now()));
  var names = (JSON.parse(settings).staff || []).map(function(x){ return x.name; }).join(', ');
  Logger.log('Настройки восстановлены. Сотрудники: ' + names);
  return 'Восстановлено. Сотрудники: ' + names;
}
function restoreSettings0912() {
  return restoreSettingsFromCopy('1Mh6IQDopBaxVg7Ek7akCQsIFz8kEujGKBMnTmqhSOuk');   // АРХИВ «CRM 2026-09-12»
}

function writeBlobs(map) {
  var sh = sheet(SVC);
  sh.clear();
  var rows = [];
  Object.keys(map).forEach(function(k) {
    var str = map[k] || '', i = 0;
    do {
      rows.push([k, i, str.substr(i * CHUNK, CHUNK)]);
      i++;
    } while (i * CHUNK < str.length);
  });
  if (rows.length) sh.getRange(1, 1, rows.length, 3).setValues(rows);
  if (!sh.isSheetHidden()) sh.hideSheet();
}

function readBlobs() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SVC);
  if (!sh || sh.getLastRow() === 0) return {};
  var vals = sh.getDataRange().getValues(), out = {};
  vals.sort(function(a, b) { return Number(a[1]) - Number(b[1]); });
  vals.forEach(function(r) {
    if (!r[0]) return;
    out[r[0]] = (out[r[0]] || '') + String(r[2] === undefined ? '' : r[2]);
  });
  return out;
}

/* ---------- мелочи ---------- */
function sheet(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}
function idx(key) {
  for (var i = 0; i < COLS.length; i++) if (COLS[i][0] === key) return i;
  return -1;
}
function hide(sh, keys) {
  keys.forEach(function(k){ sh.hideColumns(idx(k) + 1); });
}
function asDate(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return v ? String(v) : '';
}
function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
