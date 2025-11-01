/**
 * Ежемесячный сброс логов (Log и DailyLog)
 * Период: с 20 числа прошлого месяца по 19 текущего
 * Запускается каждый день в 00:10 по триггеру time-based
 */
function autoResetLogs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName("Logs");
  const dailySheet = ss.getSheetByName("DailyLog");

  const today = new Date();
  const day = today.getDate();

  // Выполняем сброс только 20-го числа месяца
  if (day !== 20) return;

  try {
    if (logSheet) {
      const last = logSheet.getLastRow();
      if (last > 1) logSheet.deleteRows(2, last - 1); // сохраняем заголовки
      else logSheet.clear(); // на всякий случай
    }

    if (dailySheet) {
      dailySheet.clear();
      const header = ["Name", "Urgent", "Non-urgent", "Total"];
      dailySheet.getRange(1, 1, 1, header.length).setValues([header]);
      dailySheet.getRange(1, 1, 1, header.length)
        .setFontWeight("bold").setFontSize(14)
        .setBackground("#FFD966")
        .setHorizontalAlignment("center");
    }

    logMonthlyReset(); // записываем в Log обнуления
  } catch (err) {
    Logger.log("❌ Ошибка при очистке логов: " + err);
  }
}

/**
 * Добавляет запись о сбросе логов в Log
 */
function logMonthlyReset() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName("Log") || ss.insertSheet("Log");
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
  logSheet.appendRow([timestamp, "SYSTEM", "Reset", "Monthly", "Logs cleared"]);
}

/**
 * Установка ежесуточного триггера проверки сброса
 */
function setupAutoResetTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "autoResetLogs") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("autoResetLogs")
    .timeBased()
    .everyDays(1)
    .atHour(0) // полночь
    .create();
}
