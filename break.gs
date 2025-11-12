/**********************
 * Автоматическая вставка брейков
 **********************/
/**********************
 * Автоматическая вставка брейков + подсветка в BreakSchedule
 **********************/
/**********************
 * Автоматическая вставка брейков + фильтрация служебных строк
 **********************/
function autoInsertBreaks() {
  const startTime = Date.now();
  const ss = SpreadsheetApp.getActiveSpreadsheet(); 
  const sheet = ss.getSheetByName("Sheet1");
  if (!sheet) return logError("❌ Лист Sheet1 не найден!");

  const lock = LockService.getScriptLock();
if (!lock.tryLock(30000)) return logError("⏳ autoInsertBreaks: занято другим процессом");


  try {
    const schedule = readBreakSchedule();
    if (!schedule.length) return logError("⚠️ Пустое расписание брейков");

    const now = new Date();
    const currentMinutes = now.getHours() * 60 + now.getMinutes();

    const currentSlot = getCurrentSlot(schedule, currentMinutes);
    const nextSlot = getNextSlot(schedule, currentSlot);

    // Подсветка активных слотов
    highlightBreakSlots(currentSlot, nextSlot);

    // Кэшированные имена
    const currentNames = new Set(currentSlot?.names || []);
    const nextNames = new Set(nextSlot?.names || []);

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    // Чтение всех имён
    const nameRange = sheet.getRange(2, 1, lastRow - 1, 1);
    const urgentRange = sheet.getRange(2, 2, lastRow - 1, 1);
    const names = nameRange.getValues().map(r => String(r[0]).trim());
    const urgentVals = urgentRange.getValues();

    const forbidden = ["name", "system"];
    const minutes = now.getMinutes();

    // Обработка в памяти
    for (let i = 0; i < names.length; i++) {
      const name = names[i];
      if (!name || forbidden.includes(name.toLowerCase())) continue;

      const currentVal = String(urgentVals[i][0] || "").toLowerCase();
      if (["sick", "vacation"].includes(currentVal)) continue;

      if (currentNames.has(name)) {
        urgentVals[i][0] = "break";
      } else if (nextNames.has(name) && minutes >= 50) {
        urgentVals[i][0] = "break coming";
      } else if (["break", "break coming"].includes(currentVal)) {
        urgentVals[i][0] = "";
      }
    }

    // Обновление листа за один вызов
    urgentRange.setValues(urgentVals);

    // Цветовая подсветка
    colorizeStatusesAndConflicts(sheet);

    const execTime = (Date.now() - startTime) / 1000;
    Logger.log(`✅ autoInsertBreaks завершён за ${execTime.toFixed(1)} сек.`);
  } catch (err) {
    logError("autoInsertBreaks: " + err);
  } finally {
    lock.releaseLock();
  }
}



/**********************
 * Подсветка активного и следующего слота
 **********************/
function highlightBreakSlots(currentSlot, nextSlot) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("BreakSchedule");
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  const range = sheet.getRange(2, 1, lastRow - 1, 2);
  const values = range.getValues();
  const backgrounds = [];
  const fonts = [];

  for (let i = 0; i < values.length; i++) {
    const [timeRange] = values[i];
    const m = String(timeRange).match(/^(\d{2}):(\d{2})-(\d{2}):(\d{2})$/);
    if (!m) {
      backgrounds.push(["white", "white"]);
      fonts.push(["normal", "normal"]);
      continue;
    }

    const start = parseInt(m[1]) * 60 + parseInt(m[2]);
    const end = parseInt(m[3]) * 60 + parseInt(m[4]);

    if (currentSlot && start === currentSlot.start && end === currentSlot.end) {
      backgrounds.push(["#C6EFCE", "#C6EFCE"]);
      fonts.push(["bold", "bold"]);
    } else if (nextSlot && start === nextSlot.start && end === nextSlot.end) {
      backgrounds.push(["#FFF2CC", "#FFF2CC"]);
      fonts.push(["bold", "bold"]);
    } else {
      backgrounds.push(["white", "white"]);
      fonts.push(["normal", "normal"]);
    }
  }

  safeSet(range, backgrounds, fonts);
}

function safeSet(range, backgrounds, fonts) {
  for (let i = 0; i < 3; i++) {
    try {
      range.setBackgrounds(backgrounds);
      range.setFontWeights(fonts);
      return;
    } catch (e) {
      Utilities.sleep(500 * (i + 1));
    }
  }
  throw new Error("safeSet: failed to update highlights");
}


/**********************
 * Создание стилизованного листа BreakSchedule
 **********************/
const CACHE_EXPIRY = 60; // секунд
function readBreakSchedule() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get("breakSchedule");
  if (cached) return JSON.parse(cached);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("BreakSchedule");
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  const result = [];

  for (let i = 0; i < data.length; i++) {
    const [timeRange, namesRaw] = data[i];
    if (!timeRange) continue;
    const m = String(timeRange).match(/^(\d{2}):(\d{2})-(\d{2}):(\d{2})$/);
    if (!m) continue;

    const start = parseInt(m[1]) * 60 + parseInt(m[2]);
    const end = parseInt(m[3]) * 60 + parseInt(m[4]);
    const names = String(namesRaw || "")
      .split(/[,|]/)
      .map(n => n.trim())
      .filter(Boolean);
    result.push({ start, end, names });
  }

  cache.put("breakSchedule", JSON.stringify(result), CACHE_EXPIRY);
  return result;
}



/**********************
 * Поиск текущего и следующего слота
 **********************/
function getCurrentSlot(schedule, currentMinutes) {
  for (const slot of schedule) {
    if (currentMinutes >= slot.start && currentMinutes < slot.end) return slot;
  }
  return schedule[schedule.length - 1];
}
function getNextSlot(schedule, currentSlot) {
  const idx = schedule.indexOf(currentSlot);
  return schedule[(idx + 1) % schedule.length];
}

/**********************
 * Вспомогательные
 **********************/
function safeSetValues(range, values, retries = 3) {
  for (let i = 0; i < retries; i++) {
    try {
      range.setValues(values);
      return;
    } catch (e) {
      Utilities.sleep(500 * (i + 1));
    }
  }
  throw new Error("safeSetValues failed");
}

function logError(msg) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = ss.getSheetByName("Logs") || ss.insertSheet("Logs");
  log.appendRow([new Date(), msg]);
  Logger.log(msg);
}
