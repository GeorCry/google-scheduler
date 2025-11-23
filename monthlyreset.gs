/**
 * Ежемесячный сброс логов (Logs, DailyLog, Log)
 * Период: с 20 числа прошлого месяца по 19 текущего.
 * Запускается каждый день в 00:10 по триггеру time-based.
 */
function autoResetLogs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logsSheet = ss.getSheetByName("Logs");   // основной лог
  const dailySheet = ss.getSheetByName("DailyLog"); 
  const logSystemSheet = ss.getSheetByName("Log"); // системный лог

  const today = new Date();
  const day = today.getDate();

  // Выполняем сброс только 20-го числа месяца
  if (day !== 20) return;

  try {
    // --- Очистка Logs (оставляем заголовок) ---
    if (logsSheet) {
      const last = logsSheet.getLastRow();
      if (last > 1) logsSheet.deleteRows(2, last - 1);
    }

    // --- Очистка DailyLog ---
    if (dailySheet) {
      dailySheet.clear();
      const header = ["Name", "Urgent", "Non-urgent", "Total"];
      dailySheet.getRange(1, 1, 1, header.length).setValues([header]);
      dailySheet.getRange(1, 1, 1, header.length)
        .setFontWeight("bold")
        .setFontSize(14)
        .setBackground("#FFD966")
        .setHorizontalAlignment("center");
    }

    // --- Очистка Log (системного), тоже оставляем заголовки ---
    if (logSystemSheet) {
      const last = logSystemSheet.getLastRow();
      if (last > 1) logSystemSheet.deleteRows(2, last - 1);
    }

    // ❗ Убрали logMonthlyReset() — больше не пишем SYSTEM записи

  } catch (err) {
    Logger.log("❌ Ошибка при очистке логов: " + err);
  }
}


/**
 * Установка ежедневного триггера
 */
function setupAutoResetTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "autoResetLogs") {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger("autoResetLogs")
    .timeBased()
    .everyDays(1)
    .atHour(0)
    .create();
}
