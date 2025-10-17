/**********************
 * Автоматическая вставка брейков
 **********************/
/**********************
 * Автоматическая вставка брейков + подсветка в BreakSchedule
 **********************/
function autoInsertBreaks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Sheet1");
  if (!sheet) return logError("❌ Лист Sheet1 не найден!");

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;

  try {
    const schedule = readBreakSchedule(); // читаем расписание (и стилизованный лист)
    if (!schedule.length) return logError("⚠️ Пустое расписание брейков");

    const now = new Date();
    const currentMinutes = now.getHours() * 60 + now.getMinutes();

    const currentSlot = getCurrentSlot(schedule, currentMinutes);
    const nextSlot = getNextSlot(schedule, currentSlot);

    // Подсветка текущего и следующего
    highlightBreakSlots(currentSlot, nextSlot);

    const currentNames = new Set(currentSlot?.names || []);
    const nextNames = new Set(nextSlot?.names || []);

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    const names = sheet.getRange(2, 1, lastRow - 1, 1).getValues().map(r => String(r[0]).trim());
    const urgentVals = sheet.getRange(2, 2, lastRow - 1, 1).getValues();

    for (let i = 0; i < names.length; i++) {
      const name = names[i];
      if (!name) continue;

      const currentVal = String(urgentVals[i][0] || "").toLowerCase();
      const isSickOrVacation = ["sick", "vacation"].includes(currentVal);
      if (isSickOrVacation) continue;

      if (currentNames.has(name)) {
        urgentVals[i][0] = "break";
      } else if (nextNames.has(name) && now.getMinutes() >= 50) {
        urgentVals[i][0] = "break coming";
      } else if (["break", "break coming"].includes(currentVal)) {
        urgentVals[i][0] = "";
      }
    }

    safeSetValues(sheet.getRange(2, 2, lastRow - 1, 1), urgentVals);
    colorizeStatusesAndConflicts(sheet);

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
  const backgrounds = range.getBackgrounds();
  const fonts = range.getFontWeights();

  // Сбрасываем всё за один проход
  for (let i = 0; i < values.length; i++) {
    const timeRange = String(values[i][0]);
    const match = timeRange.match(/^(\d{2}):(\d{2})-(\d{2}):(\d{2})$/);
    if (!match) continue;

    const start = parseInt(match[1]) * 60 + parseInt(match[2]);
    const end = parseInt(match[3]) * 60 + parseInt(match[4]);

    if (currentSlot && start === currentSlot.start && end === currentSlot.end) {
      backgrounds[i][0] = "#C6EFCE";
      backgrounds[i][1] = "#C6EFCE";
      fonts[i][0] = "bold";
      fonts[i][1] = "bold";
    } else if (nextSlot && start === nextSlot.start && end === nextSlot.end) {
      backgrounds[i][0] = "#FFF2CC";
      backgrounds[i][1] = "#FFF2CC";
      fonts[i][0] = "bold";
      fonts[i][1] = "bold";
    } else {
      backgrounds[i][0] = null;
      backgrounds[i][1] = null;
      fonts[i][0] = "normal";
      fonts[i][1] = "normal";
    }
  }

  range.setBackgrounds(backgrounds);
  range.setFontWeights(fonts);
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

