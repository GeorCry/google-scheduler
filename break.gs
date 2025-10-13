function autoInsertBreaks() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");
  if (!sheet) {
    Logger.log("Лист с таким именем не найден!");
    return;
  }

  const breakSchedule = {
    "02": ["@g.kraynik"],
    "03": ["@m.marinina"],
    "04": ["@s.denisov", "@m.poryvay"],
    "05": ["@v.lisovskaya", "@k.vagabova"],
    "06": ["@r.gabibov"]
  };

  const now = new Date();
  const currentHour = ("0" + now.getHours()).slice(-2);
  const currentMinute = now.getMinutes();
  const nextBreakHour = ("0" + ((now.getHours() + 1) % 24)).slice(-2);

  const lastRow = sheet.getLastRow();
  if (lastRow < 1) return;

  // Получаем имена и значения B колонок одним запросом
  const nameList = sheet.getRange(1, 1, lastRow, 1).getValues().map(r => r[0] ? r[0].toString().trim() : "");
  const bValues = sheet.getRange(1, 2, lastRow, 1).getValues();

  // Множества для ускорения includes
  const onBreak = new Set((breakSchedule[currentHour] || []).map(n => n.trim()));
  const coming = new Set((breakSchedule[nextBreakHour] || []).map(n => n.trim()));

  // Меняем значения в памяти
  for (let i = 0; i < nameList.length; i++) {
    const name = nameList[i];
    if (!name) continue;

    if (onBreak.has(name)) {
      bValues[i][0] = "break";
    } else if (currentMinute >= 50 && coming.has(name) && !onBreak.has(name)) {
      bValues[i][0] = "break coming";
    } else if (bValues[i][0] === "break" || bValues[i][0] === "break coming") {
      bValues[i][0] = ""; // очищаем только если это наши статусы
    }
  }

  // Записываем одним вызовом
  sheet.getRange(1, 2, lastRow, 1).setValues(bValues);

  // Обновляем стили
  colorizeStatusesAndConflicts(sheet);
}
