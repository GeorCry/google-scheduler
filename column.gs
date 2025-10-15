/***** Константы *****/
const PROP_MUTE     = 'MUTE_EDIT';
const PROP_LAST_LOG = 'LAST_LOG';

/***** Основной обработчик правок *****/
function handleEdit(e) {
  try {
    if (!e || !e.range) {
      Logger.log("⚠️ handleEdit вызван вручную");
      return;
    }

    const range = e.range;
    const sheet = range.getSheet();
    const sheetName = sheet.getName();
    if (sheetName !== "Sheet1") return;

    const col = range.getColumn();
    const row = range.getRow();
    const props = PropertiesService.getDocumentProperties();
    const lock = LockService.getDocumentLock();
    lock.waitLock(5000);

    const editor = Session.getEffectiveUser().getEmail();
    const timestampVal = timestamp();

    const type =
      col === 2 ? "Urgent" :
      col === 3 ? "Non-urgent" :
      null;

    // Только если колонка релевантная
    if (!type) return;

    const newValue = e.value || "";
    const oldValue = e.oldValue || "";
    const lowerNew = String(newValue).toLowerCase();
    const lowerOld = String(oldValue).toLowerCase();

    const isStatus = (val) => ["sick", "vacation", "break"].includes(String(val).toLowerCase());

    const isRelevantChange =
      newValue !== oldValue ||
      (newValue === "" && oldValue);

    if (!isRelevantChange) return;

    // 🟢 Логирование
    const key = [sheetName, row, col, type, newValue, editor].join("|");
    if (props.getProperty(PROP_LAST_LOG) !== key) {
      appendLogRow([timestampVal, sheet.getRange(row, 1).getValue(), sheetName, type, newValue, editor]);
      props.setProperty(PROP_LAST_LOG, key);
    }

    // 🧩 Проверяем, нужно ли пересобирать очередь
    const shouldRebuild =
      isStatus(lowerNew) ||
      (lowerOld && isStatus(lowerOld)) ||
      (!newValue && isStatus(lowerOld));

    mute(() => {
  // 1️⃣ Если все ячейки в колонке заполнены — очищаем "1" и логируем
  const cleared = clearOnesIfAllFilled(sheet, col);
  if (cleared) {
    appendLogRow([
      timestampVal,
      sheet.getRange(row, 1).getValue(),
      sheetName,
      `Clear ${type}`,
      "",
      editor
    ]);
  }

  // 2️⃣ Лёгкое оформление — без тяжёлых операций
  applyStylesSafe(sheet);

  // 3️⃣ Если изменён статус — сначала пересобираем очередь duty
  if (shouldRebuild) {
    Logger.log("♻️ Изменение статусов — пересчитываем очередь duty");
    rebuildAndApplyDuty();
  }

  // 4️⃣ Только теперь применяем цвета (чтобы Duty был актуальный)
  colorizeStatusesAndConflicts(sheet);

  // 5️⃣ Быстрая вставка брейков
  autoInsertBreaks();
});




    exportMonthStats();

  } catch (err) {
    Logger.log("❌ Ошибка handleEdit: " + err);
  }
}





/***** Хелперы *****/
function timestamp() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
}

function mute(fn) {
  const props = PropertiesService.getDocumentProperties();
  props.setProperty(PROP_MUTE, '1');
  try { fn(); } finally { props.deleteProperty(PROP_MUTE); }
}

function appendLogRow(row) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = ss.getSheetByName('Log');
  log.appendRow(row);
}

function clearOnesIfAllFilled(sheet, col) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;

  const names = sheet.getRange(2, 1, lastRow - 1, 1).getValues().map(r => String(r[0]).trim());
  const vals  = sheet.getRange(2, col, lastRow - 1, 1).getValues().map(r => r[0]);
  const dutyVals = sheet.getRange(2, 4, lastRow - 1, 1).getValues().map(r => String(r[0]).trim().toLowerCase());

  const rowIds = [];
  for (let i = 0; i < names.length; i++) if (names[i] !== "") rowIds.push(i);
  if (!rowIds.length) return false;

  const isStatus = (v, colNum) => {
    const s = String(v || "").trim().toLowerCase();
    if (!s) return false;
    if (colNum === 4) return ["pause", "queue"].includes(s) || s.includes("break");
    return ["pause", "queue", "sick", "vacation"].includes(s) || s.includes("break");
  };

  const isDutyStatus = (v) => {
    const s = String(v || "").trim().toLowerCase();
    return s === "on duty" || s === "duty coming";
  };

  const allFilled = rowIds.every(i => {
    const cellVal = String(vals[i] || "").trim();
    const dutyVal = dutyVals[i] || "";

    if (col === 2) return cellVal === "1" || isStatus(cellVal, col) || isDutyStatus(dutyVal);
    if (col === 3) return cellVal === "1" || isStatus(cellVal, col);
    if (col === 4) return cellVal === "1" || isDutyStatus(cellVal);
    return false;
  });

  if (!allFilled) return false;

  let changed = false;
  for (let i of rowIds) {
    if (String(vals[i] || "").trim() === "1") {
      vals[i] = "";
      changed = true;
    }
  }
  if (!changed) return false;

  sheet.getRange(2, col, lastRow - 1, 1).setValues(vals.map(v => [v]));
  return true;
}

function applyStylesSafe(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  sheet.setColumnWidths(1, 3, 130);
  sheet.setColumnWidth(1, 170);

  sheet.getRange(2, 1, lastRow - 1, 1)
    .setHorizontalAlignment("left").setFontWeight("bold").setFontSize(11);
  sheet.getRange(2, 2, lastRow - 1, 4)
    .setHorizontalAlignment("center").setFontWeight("bold").setFontSize(11);
}


/***** Стили *****/
function applyStylesSafe(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  sheet.setColumnWidths(1, 3, 130);
  sheet.setColumnWidth(1, 170);

  sheet.getRange(2, 1, lastRow - 1, 1)
    .setHorizontalAlignment("left").setFontWeight("bold").setFontSize(11);
  sheet.getRange(2, 2, lastRow - 1, 4)
    .setHorizontalAlignment("center").setFontWeight("bold").setFontSize(11);
}

/***** Настройка единственного триггера *****/
function setupSingleOnEditTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getEventType() === ScriptApp.EventType.ON_EDIT) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('handleEdit')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();
}

/***** Сброс глушилки *****/
function resetMuteFlag() {
  PropertiesService.getDocumentProperties().deleteProperty(PROP_MUTE);
}

/***** Экспорт месячного лога *****/
function exportMonthStats() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName("Log");
  if (!logSheet) throw new Error('Лист "Log" не найден');

  // Гарантируем наличие листа DailyLog
  let daily = ss.getSheetByName("DailyLog");
  if (!daily) daily = ss.insertSheet("DailyLog");

  // Полная очистка перед записью
  daily.clear();

  // Шапка
  const header = ["Name", "Urgent", "Non-urgent", "Total"];
  daily.getRange(1, 1, 1, header.length).setValues([header]);

  // Читаем данные Log (колонки: Timestamp, Name, Sheet, Type, Value, Edited by)
  const last = logSheet.getLastRow();
  const data = last > 1 ? logSheet.getRange(2, 1, last - 1, 6).getValues() : [];

  // Нормализация имён (и маппинг ник/почта -> логин)
  const nameMap = {
    "@a.prigozhiy": "a.prigozhiy", "a.prigozhiy@emergingtravel.com": "a.prigozhiy",
    "@milana.marinina": "m.marinina", "m.marinina@emergingtravel.com": "m.marinina",
    "@v.lisovskaya": "v.lisovskaya", "v.lisovskaya@emergingtravel.com": "v.lisovskaya",
    "@m.poryvay": "m.poryvay", "m.poryvay@emergingtravel.com": "m.poryvay",
    "@g.kraynik": "g.kraynik", "g.kraynik@emergingtravel.com": "g.kraynik",
    "@k.vagabova": "k.vagabova", "k.vagabova@emergingtravel.com": "k.vagabova",
    "@r.gabibov": "r.gabibov", "r.gabibov@emergingtravel.com": "r.gabibov",
    "@stepan.denisov": "stepan.denisov", "stepan.denisov@emergingtravel.com": "s.denisov"
  };

  const normalizeName = (val) => {
    if (!val) return "";
    let s = String(val).trim();
    if (nameMap[s]) return nameMap[s];        // точное совпадение
    s = s.toLowerCase();
    if (nameMap[s]) return nameMap[s];        // точное в lower
    if (s.startsWith("@")) s = s.slice(1);    // убираем @
    if (s.includes("@")) s = s.split("@")[0]; // почта -> логин до @
    return s;
  };

  // Подсчёт
  const stats = {};
  data.forEach(row => {
    const rawName = row[1];
    const type = row[3]; // "Urgent" | "Non-urgent" | др.
    const name = normalizeName(rawName);
    if (!name) return;

    if (!stats[name]) stats[name] = { u: 0, n: 0 };
    if (type === "Urgent") stats[name].u++;
    else if (type === "Non-urgent") stats[name].n++;
  });

  // Формируем строки + итоги
  const lines = Object.keys(stats)
    .sort((a, b) => a.localeCompare(b, "ru"))
    .map(n => [n, stats[n].u, stats[n].n, stats[n].u + stats[n].n]);

  let cursor = 2;
  if (lines.length) {
    daily.getRange(cursor, 1, lines.length, 4).setValues(lines);
    cursor += lines.length;
  }

  const totalUrgent = lines.reduce((s, r) => s + r[1], 0);
  const totalNonUrgent = lines.reduce((s, r) => s + r[2], 0);
  const totalsRow = ["TOTAL", totalUrgent, totalNonUrgent, totalUrgent + totalNonUrgent];

  daily.getRange(cursor, 1, 1, 4).setValues([totalsRow]);

  // Оформление
  daily.setFrozenRows(1);
  daily.setColumnWidths(1, 1, 180); // Name
  daily.setColumnWidths(2, 3, 120); // цифры

  // Общий шрифт чуть крупнее
  daily.getRange(1, 1, cursor, 4).setFontSize(12);

  // Шапка жирным, яркая
  daily.getRange(1, 1, 1, 4)
    .setFontWeight("bold").setFontSize(14)
    .setBackground("#FFD966")
    .setHorizontalAlignment("center");

  // Имена жирным слева
  if (lines.length) {
    daily.getRange(2, 1, lines.length, 1)
      .setFontWeight("bold")
      .setHorizontalAlignment("left");
  }

  // Числа по центру
  daily.getRange(2, 2, Math.max(1, lines.length + 1), 3).setHorizontalAlignment("center");

  // Итого внизу зелёным и жирным
  daily.getRange(cursor, 1, 1, 4)
    .setFontWeight("bold")
    .setBackground("#C6E0B4")
    .setHorizontalAlignment("center");

  // Полосатая заливка строк (без шапки)
  if (cursor > 2) {
    daily.getRange(2, 1, cursor - 2, 4).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
  }
}

function colorizeStatusesAndConflicts(sheet) {
  fixHeader(sheet);
  fixTableBorders(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  // определяем реально заполненные строки (по колонке Name = A)
  const names = sheet.getRange(2, 1, lastRow - 1, 1).getValues().map(r => String(r[0]).trim());
  let activeRows = names.findLastIndex(v => v !== "") + 1; // номер последней непустой строки
  if (activeRows < 1) return;

  // --- 2. Данные ---
  const breakCol   = 2; // Urgent
  const nonUrgCol  = 3; // Non-urgent
  const dutyCol    = 4; // Duty

  const breakVals  = sheet.getRange(2, breakCol, activeRows, 1).getValues();
  const nonUrgVals = sheet.getRange(2, nonUrgCol, activeRows, 1).getValues();
  const dutyVals   = sheet.getRange(2, dutyCol, activeRows, 1).getValues();

  const breakColors   = [];
  const nonUrgColors  = [];
  const dutyColors    = [];

  // --- 3. Цвета по статусам ---
  for (let i = 0; i < activeRows; i++) {
    const breakVal   = String(breakVals[i][0]).trim().toLowerCase();
    const nonUrgVal  = String(nonUrgVals[i][0]).trim().toLowerCase();
    const dutyVal    = String(dutyVals[i][0]).trim().toLowerCase();

    let bColor = null;
    let nColor = null;
    let dColor = null;

    // Urgent
if (breakVal === "break") {
  bColor = "#b32400";
} else if (breakVal === "break coming") {
  bColor = "#d60400";
} else if (breakVal === "pause") {
  bColor = "#999999";
} else if (breakVal === "sick") {
  bColor = "#3366cc"; // синий
} else if (breakVal === "vacation") {
  bColor = "#ff9900"; // оранжевый
}

// Non-urgent
if (nonUrgVal === "pause") {
  nColor = "#999999";
} else if (nonUrgVal === "sick") {
  nColor = "#3366cc";
} else if (nonUrgVal === "vacation") {
  nColor = "#ff9900";
}

// Duty — sick/vacation не обрабатываем
if (dutyVal === "on duty") {
  dColor = "#009900";
} else if (dutyVal === "duty coming") {
  dColor = "#00e300";
}



    breakColors.push([bColor]);
    nonUrgColors.push([nColor]);
    dutyColors.push([dColor]);
  }

  // --- 4. Применение цветов только до последней заполненной строки ---
  sheet.getRange(2, breakCol, activeRows, 1).setBackgrounds(breakColors);
  sheet.getRange(2, nonUrgCol, activeRows, 1).setBackgrounds(nonUrgColors);
  sheet.getRange(2, dutyCol,   activeRows, 1).setBackgrounds(dutyColors);
}


/***** Отправка письма с PDF *****/
function sendDailyLogByMail() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName("DailyLog");

  if (!logSheet) {
    throw new Error('Лист "DailyLog" не найден. Проверь имя вкладки.');
  }

  const url  = ss.getUrl().replace(/edit$/, "");
  const gid  = logSheet.getSheetId();

  const pdf  = UrlFetchApp.fetch(
    `${url}export?format=pdf&gid=${gid}&size=A4&portrait=true&fitw=true&gridlines=true`,
    { headers: { Authorization: `Bearer ${ScriptApp.getOAuthToken()}` } }
  ).getBlob().setName("DailyLog.pdf");

  MailApp.sendEmail({
    to:      "g.kraynik@emergingtravel.com,artyom.prohorenko@emergingtravel.com",
    subject: "Night Log Report",
    body:    "Доброе утро!\nВаш отчёт во вложении.",
    attachments: [pdf]
  });
}

function fixHeader(sheet) {
  const headerRange = sheet.getRange(1, 1, 1, 4); // первые 4 колонки, строка 1

  // Заголовки (можно убрать если не нужно менять текст)
  const headers = ["Name", "Urgent", "Non-urgent", "Duty"];
  headerRange.setValues([headers]);

  // Формат: жирный, по центру
  headerRange.setFontWeight("bold")
             .setHorizontalAlignment("center");

  // Фон (по желанию)
  headerRange.setBackground("#0035b3"); // светло-синий например

  // Толстая синяя линия снизу
  headerRange.setBorder(
    false, false, true, false, false, false,   // только низ
    "#010100",                                // цвет линии
    SpreadsheetApp.BorderStyle.SOLID_MEDIUM   // толщина
  );
}

function fixTableBorders(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) return;

  const lastCol = 4; // фиксируем только до Duty (колонка D)
  const range = sheet.getRange(1, 1, lastRow, lastCol);

  range.setBorder(
    true, true, true, true, true, true,
    "#010102", // все стороны + внутренние линии
    SpreadsheetApp.BorderStyle.SOLID // стиль
  );
}

function sender() {
  exportMonthStats();
  sendDailyLogByMail();
}

