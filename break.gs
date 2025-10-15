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

  // Сбрасываем стиль
  range.setBackground(null).setFontWeight("normal");

  const values = range.getValues();

  for (let i = 0; i < values.length; i++) {
    const timeRange = values[i][0];
    const match = String(timeRange).match(/^(\d{2}):(\d{2})-(\d{2}):(\d{2})$/);
    if (!match) continue;

    const start = parseInt(match[1]) * 60 + parseInt(match[2]);
    const end = parseInt(match[3]) * 60 + parseInt(match[4]);

    if (currentSlot && start === currentSlot.start && end === currentSlot.end) {
      sheet.getRange(i + 2, 1, 1, 2)
        .setBackground("#C6EFCE") // ярко-зелёный
        .setFontWeight("bold");
    }

    if (nextSlot && start === nextSlot.start && end === nextSlot.end) {
      sheet.getRange(i + 2, 1, 1, 2)
        .setBackground("#FFF2CC") // светло-жёлтый
        .setFontWeight("bold");
    }
  }
}

/**********************
 * Создание стилизованного листа BreakSchedule
 **********************/
function readBreakSchedule() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("BreakSchedule");

  if (!sheet) {
    sheet = ss.insertSheet("BreakSchedule");
    sheet.getRange("A1:B1")
      .setValues([["🕒 Time Range", "👥 Names"]])
      .setFontWeight("bold")
      .setFontSize(12)
      .setBackground("#FFD966")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");

    const timeSlots = [];
    for (let h = 0; h < 24; h++) {
      for (let m = 0; m < 60; m += 30) {
        const hh = ("0" + h).slice(-2);
        const mm = ("0" + m).slice(-2);
        const endH = ("0" + ((h + (m === 30 ? 1 : 0)) % 24)).slice(-2);
        const endM = m === 30 ? "00" : "30";
        timeSlots.push([`${hh}:${mm}-${endH}:${endM}`, ""]);
      }
    }

    sheet.getRange(2, 1, timeSlots.length, 2).setValues(timeSlots);

    const examples = [
      ["02:00-03:00", "@g.kraynik"],
      ["04:00-05:00", "@stepan.denisov, @m.poryvay"],
      ["05:00-06:00", "@v.nasirov, @k.vagabova"],
      ["06:00-07:00", "@r.gabibov"]
    ];
    for (const [time, names] of examples) {
      const range = sheet.createTextFinder(time).findNext();
      if (range) sheet.getRange(range.getRow(), 2).setValue(names);
    }

    sheet.setColumnWidths(1, 1, 130);
    sheet.setColumnWidths(2, 1, 280);

    sheet.getRange("A:B")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setFontSize(11);

    sheet.getRange(1, 1, timeSlots.length + 1, 2)
      .setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);

    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, timeSlots.length + 1, 2).createFilter();

    SpreadsheetApp.flush();
    return [];
  }

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
