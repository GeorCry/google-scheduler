/**********************
 * duty.gs — Автоматическое построение ночной очереди и обновление статусов
 **********************/

const DUTY_HOURS = Array.from({length: 24}, (_, i) => String(i).padStart(2, "0"));

/* -------------------------
   Утилиты
   ------------------------ */
function normalizeName(raw) {
  // убираем неразрывные пробелы, обрезаем, снижаем регистр, убираем ведущие @
  return String(raw || "")
    .replace(/\u00A0/g, " ")
    .trim()
    .replace(/^@+/, "")
    .toLowerCase();
}

/* -------------------------
   0. Проверка/создание листов
   ------------------------ */
function ensureDutySheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let schedule = ss.getSheetByName("DutySchedule");
  if (!schedule) {
    schedule = ss.insertSheet("DutySchedule");
    const data = [["Hour", "Name", "Manual"]];

    // создаём строки для 24 часов
    for (let h = 0; h < 24; h++) {
      const hh = String(h).padStart(2, "0");
      data.push([hh, "", false]);
    }

    schedule.getRange(1,1,data.length,3).setValues(data);
    schedule.setColumnWidths(1,3,130);
    schedule.getRange("A1:C1")
      .setFontWeight("bold")
      .setBackground("#ffd966")
      .setHorizontalAlignment("center");
    schedule.getRange("A2:A25").setHorizontalAlignment("center");
  }

  let queue = ss.getSheetByName("DutyQueue");
  if (!queue) {
    queue = ss.insertSheet("DutyQueue");
    const data = [
      ["Order","Name","Manual","Active"],
      [1,"@r.gabibov",false,true],
      [2,"@g.kraynik",false,true],
      [3,"@m.marinina",false,true],
      [4,"@v.lisovskaya",false,true],
      [5,"@k.vagabova",false,true],
      [6,"@m.poryvay",false,true],
      [7,"@s.denisov",false,true],
    ];
    queue.getRange(1,1,data.length,4).setValues(data);
    queue.setColumnWidths(1,4,130);
    queue.getRange("A1:D1")
      .setFontWeight("bold")
      .setBackground("#c6e0b4")
      .setHorizontalAlignment("center");
  }

  let main = ss.getSheetByName("Sheet1");
  if (!main) {
    main = ss.insertSheet("Sheet1");
    main.getRange("A1:D1")
      .setValues([["Name","Urgent","NonUrgent","DutyStatus"]])
      .setFontWeight("bold")
      .setBackground("#d9e1f2")
      .setHorizontalAlignment("center");
    main.setColumnWidths(1,4,130);
  }

  SpreadsheetApp.flush();
  Utilities.sleep(150);
}


/* -------------------------
   1. Чтение данных
   ------------------------ */
function getDutySchedule() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DutySchedule");
  if (!sheet) return {};
  const rows = Math.max(sheet.getLastRow() - 1, 0);
  if (rows === 0) return {};
  const data = sheet.getRange(2,1,rows,3).getValues();
  const res = {};
  for (const [hour, name, manual] of data) {
  if (hour === null || hour === undefined || hour === "") continue;
  const h = String(hour).padStart(2, "0");
  res[h] = { name: name || "", manual: manual === true };
  // дублируем для числового ключа, чтобы находилось и как "0"
  res[String(Number(hour))] = res[h];
}
  return res;
}

function getDutyQueue() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DutyQueue");
  if (!sheet) return [];

  const rows = Math.max(sheet.getLastRow() - 1, 0);
  if (rows === 0) {
    // fallback: собрать из DutySchedule (без лишних дубликатов)
    const sched = getDutySchedule();
    const unique = [...new Set(Object.values(sched).map(x => x.name).filter(Boolean))];
    return unique.map((n,i) => ({ order: i+1, name: String(n).trim(), manualFlag:false, active:true }));
  }

  const data = sheet.getRange(2,1,rows,4).getValues();
  return data
    .filter(r => r[1] && String(r[1]).trim() !== "")
    .sort((a,b) => (Number(a[0])||0) - (Number(b[0])||0))
    .map(r => ({ order: r[0], name: String(r[1]).trim(), manualFlag: r[2] === true, active: r[3] === true }));
}

function getUnavailable() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Sheet1");
  if (!sheet) return new Set();

  // Гарантируем, что все последние правки применены
  SpreadsheetApp.flush();
  Utilities.sleep(100);

  const rows = Math.max(sheet.getLastRow() - 1, 0);
  if (rows === 0) return new Set();

  const values = sheet.getRange(2,1,rows,3).getValues();
  const set = new Set();

  for (const [nameRaw, urgentRaw, nonUrgRaw] of values) {
    const name = String(nameRaw || "").trim();
    if (!name) continue;

    const urgent = String(urgentRaw || "").toLowerCase();
    const nonUrg = String(nonUrgRaw || "").toLowerCase();

    // ищем вхождение ключевых слов (includes) — чтобы ловить "vacation until ..." и пр.
    if (["sick","vacation","break"].some(k => (urgent.indexOf(k) !== -1) || (nonUrg.indexOf(k) !== -1))) {
      // добавляем в множестве как нормализованное имя (и с @-версией), чтобы сравнения всегда работали
      const nNorm = normalizeName(name);
      set.add(nNorm);
      set.add("@" + nNorm);
    }
  }

  return set;
}

/* -------------------------
   2. Авто-построение
   ------------------------ */
function autoDutyScheduler() {
  ensureDutySheets();

  const queue = getDutyQueue();
  if (!queue.length) {
    Logger.log("⚠️ Очередь пуста, остановка");
    SpreadsheetApp.getActiveSpreadsheet().toast("⚠️ Очередь пуста", "Duty Scheduler", 3);
    return {};
  }

  const unavailable = getUnavailable();
  const current = getDutySchedule();

  // 🧭 Фильтруем только активных, не в отпуске/больничном
  let pool = queue
    .filter(q => q.active && q.name && !unavailable.has(normalizeName(q.name)))
    .map(q => q.name);

  // fallback — если все недоступны, берём всех активных
  if (!pool.length) {
    pool = queue.filter(q => q.active && q.name).map(q => q.name);
  }

  if (!pool.length) {
    Logger.log("⚠️ Нет доступных участников");
    SpreadsheetApp.getActiveSpreadsheet().toast("⚠️ Нет активных участников", "Duty Scheduler", 3);
    return {};
  }

  // 📊 Подсчёт уже назначенных (исторически)
  const pastStats = {};
  for (const h of Object.keys(current)) {
    const entry = current[h];
    if (entry && entry.name) {
      const n = normalizeName(entry.name);
      pastStats[n] = (pastStats[n] || 0) + 1;
    }
  }

  // Инициализация счётчиков
  const assignCounts = Object.fromEntries(
    pool.map(p => [normalizeName(p), pastStats[normalizeName(p)] || 0])
  );

  const result = {};
  let prev = null;

  // ⚙️ Гарантируем, что ключи будут "00", "01", ..., "23"
  const hours = DUTY_HOURS.map(h => String(h).padStart(2, "0"));

  for (const h of hours) {
    const existing = current[h];

    // 🔒 Пропускаем ручные назначения и прочерки
if (existing && existing.name) {
  const nameStr = String(existing.name).trim();
  // если вручную стоит "-", не трогаем этот час
  if (nameStr === "-") {
    result[h] = { name: "-", manual: true };
    prev = null; // не учитываем как предыдущего дежурного
    continue;
  }
  if (existing.manual === true || existing.manual === "TRUE") {
    result[h] = { name: nameStr, manual: true };
    const n = normalizeName(nameStr);
    if (assignCounts.hasOwnProperty(n)) assignCounts[n]++;
    prev = nameStr;
    continue;
  }
}


    // 🎯 Кандидаты — все, кроме предыдущего
    const candidates = pool.filter(p => normalizeName(p) !== normalizeName(prev));

    // ⚖️ Выбор с наименьшим числом назначений
    candidates.sort((a, b) => {
      const diff =
        (assignCounts[normalizeName(a)] || 0) - (assignCounts[normalizeName(b)] || 0);
      if (diff !== 0) return diff;
      return Math.random() - 0.5; // случайный выбор при равенстве
    });

    const chosen = candidates[0] || pool[0];
    result[h] = { name: chosen, manual: false };
    assignCounts[normalizeName(chosen)] =
      (assignCounts[normalizeName(chosen)] || 0) + 1;
    prev = chosen;
  }

  Logger.log("🕒 Сформировано расписание по часам: " + Object.keys(result).join(", "));

  writeDutySchedule(result);
  SpreadsheetApp.getActiveSpreadsheet().toast("✅ Автопланирование завершено", "Duty Scheduler", 3);
  Logger.log("✅ Автопланирование завершено (равномерное распределение)");

  return result;
}

/* -------------------------
   3. Запись расписания
   ------------------------ */
function writeDutySchedule(schedule) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DutySchedule");
  if (!sheet) return;

  const out = [["Hour", "Name", "Manual"]];

  for (const rawH of DUTY_HOURS) {
    // Приводим все ключи к двухзначному формату (чтобы не было "0" вместо "00")
    const h = String(rawH).padStart(2, "0");
    const item = schedule[h] || schedule[String(Number(h))] || { name: "", manual: false };
    out.push([h, item.name, item.manual]);
  }

  // перезаписываем таблицу
  sheet.clearContents();
  sheet.getRange(1, 1, out.length, 3).setValues(out);
  sheet.setColumnWidths(1, 3, 120);
  sheet.getRange("A1:C1").setFontWeight("bold").setBackground("#ffd966");

  SpreadsheetApp.flush();
  Logger.log("✅ DutySchedule обновлён (часы нормализованы)");
}


/* -------------------------
   4. Обновление статусов на основном листе
   ------------------------ */
/* -------------------------
   4. Обновление статусов на основном листе
   ------------------------ */
function updateDutyStatus() {
  ensureDutySheets();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Sheet1");
  if (!sheet) return;

  let duty = getDutySchedule();
  if (!Object.keys(duty).length) duty = autoDutyScheduler();

  // вспом. функция: найдём имя дежурного и для "00"/"0" вариантов
  function findDutyNameForHour(dutyObj, hourStr) {
    const hStr = String(hourStr).padStart(2, "0"); // "00"
    const hNum = String(Number(hourStr));          // "0"
    return (dutyObj[hStr] && dutyObj[hStr].name) || (dutyObj[hNum] && dutyObj[hNum].name) || "";
  }

  // если в расписании стоит "-", то считаем этот слот пропущенным
  function isHourSkipped(name) {
    return String(name || "").trim() === "-";
  }

  const now = new Date();
  const tz = Session.getScriptTimeZone() || "GMT+3";
  const curH = Utilities.formatDate(now, tz, "HH");
  const nextH = Utilities.formatDate(new Date(now.getTime() + 3600000), tz, "HH");
  const curM = now.getMinutes();

  let onDutyNameRaw = findDutyNameForHour(duty, curH);
  let comingNameRaw = findDutyNameForHour(duty, nextH);

  // Если для текущего часа стоит прочерк — просто делаем так,
  // чтобы дальше логика не ставила "on duty" никому.
  if (isHourSkipped(onDutyNameRaw)) {
    onDutyNameRaw = "";
    Logger.log(`⏭ Пропуск часа ${curH}:00 — стоит прочерк`);
  }
  // Для следующего часа — если стоит прочерк, не ставим "duty coming"
  if (isHourSkipped(comingNameRaw)) {
    comingNameRaw = "";
    Logger.log(`⏭ Пропуск следующего часа ${nextH}:00 — стоит прочерк`);
  }

  const onDutyNorm = normalizeName(onDutyNameRaw);
  const comingNorm = normalizeName(comingNameRaw);

  const total = Math.max(sheet.getLastRow() - 1, 0);
  if (total === 0) return;

  const names = sheet.getRange(2, 1, total, 1).getValues().map(r => String(r[0] || "").trim());
  const statuses = sheet.getRange(2, 4, total, 1).getValues();

  for (let i = 0; i < names.length; i++) {
    const nameNorm = normalizeName(names[i]);
    if (nameNorm && onDutyNorm && nameNorm === onDutyNorm) {
      statuses[i][0] = "on duty";
    } else if (nameNorm && comingNorm && nameNorm === comingNorm && curM >= 50) {
      statuses[i][0] = "duty coming";
    } else {
      // очищаем статус (если не ручной — текущая логика очищает всегда)
      statuses[i][0] = "";
    }
  }

  sheet.getRange(2, 4, total, 1).setValues(statuses);
  SpreadsheetApp.flush();
  Logger.log("✅ Статусы на основном листе обновлены");
}



/* -------------------------
   5. Главная функция — обновление Active и расписания
   ------------------------ */
function rebuildAndApplyDuty() {
  ensureDutySheets();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const queueSheet = ss.getSheetByName("DutyQueue");
  const dutySheet = ss.getSheetByName("DutySchedule");

  // Принудительно применим последние правки
  SpreadsheetApp.flush();
  Utilities.sleep(150);

  const unavailable = getUnavailable();
  // === 1️⃣ Проверка доступности участников очереди ===
  if (queueSheet) {
    const rows = Math.max(queueSheet.getLastRow() - 1, 0);
    if (rows > 0) {
      const range = queueSheet.getRange(2, 1, rows, 4);
      const data = range.getValues();
      let changed = false;

      for (let i = 0; i < data.length; i++) {
  const rawHour = data[i][0];
  const rawName = String(data[i][1] || "").trim();
  const nameCell = data[i][1];

  if (!rawHour || !nameCell || rawName === "-") continue;

  const nNorm = normalizeName(rawName);
  const wasActive = data[i][3] === true;
  const nowActive = !unavailable.has(nNorm) && !unavailable.has("@" + nNorm);

  if (wasActive !== nowActive) {
    data[i][3] = nowActive;
    changed = true;

    const msg = nowActive
      ? `${rawName} снова активен`
      : `${rawName} временно исключён (sick/vacation)`;

    ss.toast(msg, "Duty Update", 3);
    Logger.log("🔔 " + msg);
  }
}


      if (changed) {
        range.setValues(data);
        SpreadsheetApp.flush();
        Utilities.sleep(150);
      }
    }
  }

  // === 2️⃣ Нормализуем ключи расписания (чтобы "0" → "00") ===
  const duty = {};
  if (dutySheet) {
    const data = dutySheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const rawHour = data[i][0];
      const nameCell = data[i][1];
      if (!rawHour || !nameCell) continue;

      const hourKey = String(rawHour).padStart(2, "0"); // "0" → "00"
      duty[hourKey] = { name: String(nameCell).trim() };
    }
  }

  // === 3️⃣ Определяем текущего и следующего дежурного ===
  const tz = Session.getScriptTimeZone() || "GMT+3";
  const now = new Date();
  const curH = Utilities.formatDate(now, tz, "HH");
  const nextH = Utilities.formatDate(new Date(now.getTime() + 3600000), tz, "HH");

  const curDuty = duty[curH]?.name || "";
  const nextDuty = duty[nextH]?.name || "";

  Logger.log(`🕛 Текущее время: ${curH}:00 — OnDuty: ${curDuty} → Next: ${nextDuty}`);

  // === 4️⃣ Пересчёт очереди и обновление Duty статусов ===
  autoDutyScheduler();
  Utilities.sleep(150);
  updateDutyStatus();

  ss.toast("✅ Очередь и дежурства обновлены", "Duty Scheduler", 3);
  Logger.log("✅ Полное обновление завершено");
}

