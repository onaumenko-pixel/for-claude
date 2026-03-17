// ================================================================
//  PIPEDRIVE → GOOGLE SHEETS  |  Авто-синхронізація щопонеділка
// ================================================================

var CONFIG = {
  API_TOKEN:      'YOUR_PIPEDRIVE_API_TOKEN',
  SPREADSHEET_ID: '1CnqOwJgtm7r_BGF4B6A9ZuDn4AziBryk0BRh6Jhlxc0',
  SHEET_NAME:     '📋 Всі відділи',
  PAGE_LIMIT:     500,
};

// ================================================================
// АВТО-ФУНКЦІЯ — кожен понеділок о 08:00 (минулий тиждень пн–нд)
// ================================================================
function syncFromPipedrive() {
  var today = new Date();
  var dow = today.getDay() === 0 ? 7 : today.getDay();
  var thisMonday = new Date(today);
  thisMonday.setDate(today.getDate() - (dow - 1));
  thisMonday.setHours(0, 0, 0, 0);
  var lastMonday = new Date(thisMonday);
  lastMonday.setDate(thisMonday.getDate() - 7);
  var lastSunday = new Date(thisMonday);
  lastSunday.setDate(thisMonday.getDate() - 1);
  lastSunday.setHours(23, 59, 59, 0);
  syncForPeriod(lastMonday, lastSunday);
}

// ================================================================
// СИНХРОНІЗАЦІЯ ЗА ВКАЗАНИЙ ПЕРІОД
// ================================================================
function syncForPeriod(dateFrom, dateTo) {
  var ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = getOrCreateSheet(ss);

  var users       = fetchUsers();
  var existingIds = getExistingIds(sheet);
  var activities  = fetchActivities(dateFrom, dateTo);
  var newRows     = [];

  activities.forEach(function(a) {
    if (existingIds.indexOf(String(a.id)) !== -1) return;
    var subject = (a.subject || '').toLowerCase();
    if (subject.indexOf('без відповіді') !== -1 || subject.indexOf('пропущено') !== -1) return;
    var doneTime = a.marked_as_done_time ? new Date(a.marked_as_done_time) : null;
    if (!doneTime || doneTime < dateFrom || doneTime > dateTo) return;

    var userName = (users[a.assigned_to_user_id] || users[a.user_id] || '').trim();

    newRows.push([
      userName,
      cleanType(a.type_name || a.type || ''),
      a.subject    || '',
      Utilities.formatDate(doneTime, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      a.org_name   || '',
      a.deal_title || '',
      a.done ? 1 : 0,
      String(a.id)
    ]);
  });

  if (newRows.length > 0) {
    var lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1, newRows.length, 8).setValues(newRows);
    sheet.hideColumns(8);
    ss.toast('✅ Додано ' + newRows.length + ' записів!', 'Готово', 5);
  } else {
    ss.toast('Нових записів немає', 'Готово', 3);
  }
}

// ================================================================
// ОТРИМАННЯ КОРИСТУВАЧІВ З PIPEDRIVE (id → name)
// ================================================================
function fetchUsers() {
  var url = 'https://api.pipedrive.com/v1/users?api_token=' + CONFIG.API_TOKEN;
  try {
    var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    var json = JSON.parse(resp.getContentText());
    var map  = {};
    if (json.success && json.data) {
      json.data.forEach(function(u) { map[u.id] = u.name; });
    }
    return map;
  } catch(e) {
    Logger.log('❌ fetchUsers: ' + e.message);
    return {};
  }
}

// ================================================================
// ОТРИМАННЯ ДАНИХ З PIPEDRIVE
// ================================================================
function fetchActivities(dateFrom, dateTo) {
  var all     = [];
  var start   = 0;
  var fromStr = Utilities.formatDate(dateFrom, 'UTC', 'yyyy-MM-dd');
  var toStr   = Utilities.formatDate(dateTo,   'UTC', 'yyyy-MM-dd');

  while (true) {
    var url = 'https://api.pipedrive.com/v1/activities'
      + '?api_token=' + CONFIG.API_TOKEN
      + '&start='      + start
      + '&limit='      + CONFIG.PAGE_LIMIT
      + '&start_date=' + fromStr
      + '&end_date='   + toStr
      + '&done=1'
      + '&user_id=0';

    try {
      var resp  = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      var json  = JSON.parse(resp.getContentText());
      if (!json.success) { Logger.log('❌ ' + resp.getContentText()); break; }
      var items = json.data || [];
      if (items.length === 0) break;
      all = all.concat(items);
      var more = json.additional_data
        && json.additional_data.pagination
        && json.additional_data.pagination.more_items_in_collection;
      if (!more) break;
      start += CONFIG.PAGE_LIMIT;
      Utilities.sleep(300);
    } catch(e) {
      Logger.log('❌ ' + e.message); break;
    }
  }
  return all;
}

// ================================================================
// ДОПОМІЖНІ ФУНКЦІЇ
// ================================================================
function getOrCreateSheet(ss) {
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
    var headers = ['Менеджер', 'Тип', 'Тема', 'Дата', 'Організація', 'Угода', 'Виконано'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
      .setBackground('#1a73e8').setFontColor('#ffffff').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getExistingIds(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 8, lastRow - 1, 1).getValues()
    .map(function(r) { return String(r[0]); })
    .filter(function(id) { return id !== ''; });
}

function cleanType(type) {
  return type.replace(/[\u{1F000}-\u{1FFFF}]/gu, '')
             .replace(/[\u{2600}-\u{27BF}]/gu, '')
             .replace(/\s+/g, ' ').trim();
}

// ================================================================
// ЗАПУСТИ ОДИН РАЗ — встановить тригер щопонеділка о 08:00
// ================================================================
function setupWeeklyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'syncFromPipedrive') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('syncFromPipedrive')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(8)
    .create();
  Logger.log('✅ Тригер встановлено: кожен понеділок о 08:00');
}
