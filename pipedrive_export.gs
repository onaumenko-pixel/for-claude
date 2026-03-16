/**
 * Pipedrive → Google Sheets Export
 * Выгрузка выполненных активностей операторов колл-центра
 *
 * Настройки: задайте значения в разделе CONFIG ниже,
 * либо через Project Settings → Script Properties (рекомендуется для токена).
 */

// ─── КОНФИГУРАЦИЯ ──────────────────────────────────────────────────────────────

const CONFIG = {
  // Токен Pipedrive API. Лучше хранить в Script Properties:
  // Project Settings → Script Properties → PIPEDRIVE_TOKEN
  PIPEDRIVE_TOKEN: PropertiesService.getScriptProperties().getProperty('PIPEDRIVE_TOKEN') || 'YOUR_PIPEDRIVE_API_TOKEN',

  // ID Google Таблицы (из URL: .../spreadsheets/d/<ID>/...)
  SPREADSHEET_ID: PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID') || 'YOUR_SPREADSHEET_ID',

  // Название вкладки для записи данных
  SHEET_NAME: 'сырые данные',

  // Базовый URL Pipedrive API
  PIPEDRIVE_BASE_URL: 'https://api.pipedrive.com/v1',

  // Лимит записей за один запрос (макс. 500)
  PAGE_LIMIT: 500,
};

// ─── СПИСОК ОПЕРАТОРОВ ─────────────────────────────────────────────────────────

/**
 * Имена операторов без приставки "Оператор".
 * Ключ — часть строки, которая будет искаться в assigned_to_user_name.
 */
const OPERATORS = [
  'Гадайчук Наталія',
  'Ходаківський Юрій',
  'Погорелова Ірина',
  'Селезньов Андрій',
  'Бобровицький Станіслав',
];

// ─── ТИПЫ ЗАДАНИЙ И БАЛЛЫ ──────────────────────────────────────────────────────

const ACTIVITY_SCORES = {
  'Дублі та інше': 1,
  'Недозвон': 2,
  'Телефонний дзвінок Клієнту': 4,
  'Надіслано лист/ КП': 5,
  'Обробка нових': 6,
  'Передача картки на МП / ПКП старий лід': 8,
  'Передача картки на МП / ПКП новий лід': 10,
};

// ─── ЗАГОЛОВКИ ТАБЛИЦЫ ─────────────────────────────────────────────────────────

const HEADERS = [
  'Тема',           // subject
  'Тип',            // type_name (очищенный)
  'Баллы',          // рассчитывается
  'Выполнено',      // done
  'Дата',           // marked_as_done_time
  'Организация',    // org_name
  'Идентификатор',  // id (для дедупликации)
  'Сделка',         // deal_title
  'Заметка',        // note (без HTML)
  'Оператор',       // assigned_to_user_name (очищенный)
];

// ─── ОСНОВНАЯ ФУНКЦИЯ ЭКСПОРТА ──────────────────────────────────────────────────

/**
 * Главная точка входа. Определяет диапазон дат и запускает экспорт.
 * Вызывается триггером каждый понедельник.
 */
function exportPreviousWeek() {
  const { startDate, endDate } = getPreviousWeekRange();
  Logger.log(`Выгрузка за период: ${startDate} — ${endDate}`);
  runExport(startDate, endDate);
}

/**
 * Ручной запуск для первой выгрузки (09.03.2026 – 15.03.2026).
 * Запустите эту функцию один раз вручную.
 */
function exportFirstWeekManual() {
  runExport('2026-03-09', '2026-03-15');
}

/**
 * Основная логика экспорта за указанный период.
 * @param {string} startDate - начало периода YYYY-MM-DD (включительно)
 * @param {string} endDate   - конец периода YYYY-MM-DD (включительно)
 */
function runExport(startDate, endDate) {
  const sheet = getOrCreateSheet();
  const existingIds = getExistingIds(sheet);

  Logger.log(`Уже существует записей: ${existingIds.size}`);

  const activities = fetchActivities(startDate, endDate);
  Logger.log(`Получено активностей из Pipedrive: ${activities.length}`);

  const rows = [];

  for (const activity of activities) {
    // Фильтр: только выполненные
    if (!activity.done) continue;

    // Очищаем имя оператора и проверяем, что он в нашем списке
    const operatorName = cleanOperatorName(activity.assigned_to_user_name || '');
    if (!isTargetOperator(operatorName)) continue;

    // Очищаем тип задания (убираем смайлы)
    const cleanType = cleanActivityType(activity.type_name || '');

    // Проверяем, что тип задания входит в разрешённый список
    const score = getScore(cleanType);
    if (score === null) continue;

    // Дедупликация по ID
    const id = String(activity.id);
    if (existingIds.has(id)) continue;

    // Формируем строку
    const row = [
      activity.subject || '',
      cleanType,
      score,
      activity.done ? 1 : 0,
      formatDate(activity.marked_as_done_time),
      activity.org_name || '',
      id,
      activity.deal_title || '',
      stripHtml(activity.note || ''),
      operatorName,
    ];

    rows.push(row);
    existingIds.add(id); // защита от дублей внутри одной выгрузки
  }

  if (rows.length === 0) {
    Logger.log('Новых записей для добавления не найдено.');
    return;
  }

  appendRows(sheet, rows);
  Logger.log(`Добавлено новых записей: ${rows.length}`);
}

// ─── РАБОТА С GOOGLE SHEETS ────────────────────────────────────────────────────

/**
 * Возвращает вкладку "сырые данные", создаёт её при необходимости.
 * При создании добавляет строку заголовков.
 */
function getOrCreateSheet() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
    sheet.appendRow(HEADERS);
    // Скрываем колонку "Идентификатор" (7-я, индекс 7)
    sheet.hideColumns(HEADERS.indexOf('Идентификатор') + 1);
    Logger.log(`Создана новая вкладка: "${CONFIG.SHEET_NAME}"`);
  }

  return sheet;
}

/**
 * Считывает все существующие ID из колонки "Идентификатор".
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @returns {Set<string>}
 */
function getExistingIds(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return new Set();

  const idColIndex = HEADERS.indexOf('Идентификатор') + 1; // 1-based
  const values = sheet.getRange(2, idColIndex, lastRow - 1, 1).getValues();
  return new Set(values.map(row => String(row[0])).filter(Boolean));
}

/**
 * Добавляет массив строк в конец таблицы.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Array<Array>} rows
 */
function appendRows(sheet, rows) {
  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, rows.length, HEADERS.length).setValues(rows);
}

// ─── РАБОТА С PIPEDRIVE API ────────────────────────────────────────────────────

/**
 * Получает все выполненные активности за указанный период.
 * Использует пагинацию для обхода лимита в 500 записей.
 * @param {string} startDate - YYYY-MM-DD
 * @param {string} endDate   - YYYY-MM-DD
 * @returns {Array<Object>}
 */
function fetchActivities(startDate, endDate) {
  const allActivities = [];
  let start = 0;
  let hasMore = true;

  // Формируем диапазон: marked_as_done_time от начала startDate до конца endDate
  const since = `${startDate} 00:00:00`;
  const until = `${endDate} 23:59:59`;

  while (hasMore) {
    const url = buildUrl('/activities', {
      done: 1,
      start: start,
      limit: CONFIG.PAGE_LIMIT,
      start_date: startDate,
      end_date: endDate,
    });

    const response = apiRequest(url);

    if (!response || !response.data) {
      Logger.log(`Нет данных на странице start=${start}`);
      break;
    }

    // Фильтруем по marked_as_done_time (API фильтрует по due_date, уточняем сами)
    const filtered = response.data.filter(activity => {
      if (!activity.marked_as_done_time) return false;
      const doneTime = activity.marked_as_done_time.substring(0, 10); // YYYY-MM-DD
      return doneTime >= startDate && doneTime <= endDate;
    });

    allActivities.push(...filtered);

    const pagination = response.additional_data && response.additional_data.pagination;
    if (pagination && pagination.more_items_in_collection) {
      start += CONFIG.PAGE_LIMIT;
    } else {
      hasMore = false;
    }
  }

  return allActivities;
}

/**
 * Выполняет GET-запрос к Pipedrive API.
 * @param {string} url
 * @returns {Object|null}
 */
function apiRequest(url) {
  try {
    const response = UrlFetchApp.fetch(url, {
      method: 'GET',
      headers: {
        'Authorization': `Bearer ${CONFIG.PIPEDRIVE_TOKEN}`,
        'Accept': 'application/json',
      },
      muteHttpExceptions: true,
    });

    const code = response.getResponseCode();
    if (code !== 200) {
      Logger.log(`Ошибка API: HTTP ${code} | ${response.getContentText().substring(0, 300)}`);
      return null;
    }

    const json = JSON.parse(response.getContentText());
    if (!json.success) {
      Logger.log(`Pipedrive вернул success=false: ${JSON.stringify(json.error)}`);
      return null;
    }

    return json;
  } catch (e) {
    Logger.log(`Исключение при запросе к API: ${e.message}`);
    return null;
  }
}

/**
 * Строит URL для Pipedrive API.
 * @param {string} path - путь, например '/activities'
 * @param {Object} params - GET-параметры
 * @returns {string}
 */
function buildUrl(path, params) {
  const base = `${CONFIG.PIPEDRIVE_BASE_URL}${path}`;
  const query = Object.entries(params)
    .map(([k, v]) => `${encodeURIComponent(k)}=${encodeURIComponent(v)}`)
    .join('&');
  return `${base}?${query}`;
}

// ─── ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ───────────────────────────────────────────────────

/**
 * Вычисляет диапазон прошлой недели (пн–вс) относительно текущего дня.
 * @returns {{ startDate: string, endDate: string }}
 */
function getPreviousWeekRange() {
  const today = new Date();
  const dayOfWeek = today.getDay(); // 0=вс, 1=пн, …, 6=сб

  // Находим прошлый понедельник
  // Если сегодня понедельник (1), то прошлый пн = -7 дней
  const daysToLastMonday = dayOfWeek === 0 ? 6 : dayOfWeek - 1;
  const lastMonday = new Date(today);
  lastMonday.setDate(today.getDate() - daysToLastMonday - 7);

  const lastSunday = new Date(lastMonday);
  lastSunday.setDate(lastMonday.getDate() + 6);

  return {
    startDate: formatDateYMD(lastMonday),
    endDate: formatDateYMD(lastSunday),
  };
}

/**
 * Форматирует дату в строку YYYY-MM-DD.
 * @param {Date} date
 * @returns {string}
 */
function formatDateYMD(date) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const d = String(date.getDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
}

/**
 * Форматирует строку datetime из Pipedrive в читаемый вид.
 * Возвращает только дату (YYYY-MM-DD), время отбрасывается согласно ТЗ.
 * @param {string} datetimeStr - например "2026-03-10 14:23:00"
 * @returns {string}
 */
function formatDate(datetimeStr) {
  if (!datetimeStr) return '';
  return datetimeStr.substring(0, 10);
}

/**
 * Очищает имя оператора: убирает приставку "Оператор" и лишние пробелы.
 * @param {string} name
 * @returns {string}
 */
function cleanOperatorName(name) {
  return name
    .replace(/^Оператор\s*/i, '')
    .trim();
}

/**
 * Проверяет, входит ли имя оператора в целевой список.
 * @param {string} cleanedName
 * @returns {boolean}
 */
function isTargetOperator(cleanedName) {
  return OPERATORS.some(op => op.toLowerCase() === cleanedName.toLowerCase());
}

/**
 * Убирает из строки типа задания эмодзи и лишние пробелы.
 * @param {string} typeName
 * @returns {string}
 */
function cleanActivityType(typeName) {
  // Убираем эмодзи (Unicode ranges для эмодзи)
  return typeName
    .replace(/[\u{1F000}-\u{1FFFF}]/gu, '')  // эмодзи из Supplementary Multilingual Plane
    .replace(/[\u{2600}-\u{27BF}]/gu, '')     // разные символы (☀, ✉, ✅ и др.)
    .replace(/[\u{FE00}-\u{FE0F}]/gu, '')     // вариативные селекторы (️)
    .replace(/\s+/g, ' ')
    .trim();
}

/**
 * Возвращает количество баллов для типа задания или null, если тип не в списке.
 * @param {string} cleanType - очищенное название типа
 * @returns {number|null}
 */
function getScore(cleanType) {
  for (const [key, score] of Object.entries(ACTIVITY_SCORES)) {
    if (key.toLowerCase() === cleanType.toLowerCase()) {
      return score;
    }
  }
  return null; // тип не входит в список — пропускаем
}

/**
 * Убирает HTML-теги из строки (для поля note).
 * @param {string} html
 * @returns {string}
 */
function stripHtml(html) {
  if (!html) return '';
  return html
    .replace(/<br\s*\/?>/gi, '\n')   // <br> → перенос строки
    .replace(/<\/p>/gi, '\n')         // </p> → перенос строки
    .replace(/<[^>]+>/g, '')          // убираем все остальные теги
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .trim();
}

// ─── НАСТРОЙКА ТРИГГЕРА ────────────────────────────────────────────────────────

/**
 * Создаёт еженедельный триггер: каждый понедельник в 09:00.
 * Запустите эту функцию ОДИН РАЗ вручную для установки триггера.
 */
function createWeeklyTrigger() {
  // Удаляем существующие триггеры на эту функцию во избежание дублирования
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'exportPreviousWeek') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Создаём новый триггер: каждый понедельник в 09:00–10:00
  ScriptApp.newTrigger('exportPreviousWeek')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(9)
    .create();

  Logger.log('Триггер создан: каждый понедельник в 09:00.');
}
