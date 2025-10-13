/**********************
 * duty.gs — Автоматическое построение ночной очереди и обновление статусов
 **********************/

const DUTY_HOURS = ["21","22","23","00","01","02","03","04","05","06","07","08"];

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
    const data = [
      ["Hour", "Name", "Manual"],
      ["21","@r.gabibov",false],
      ["22","@g.kraynik",false],
      ["23","@m.marinina",false],
      ["00","@v.lisovskaya",false],
      ["01","@k.vagabova",false],
      ["02","@s.denisov",false],
      ["03","@v.lisovskaya",false],
      ["04","@k.vagabova",false],
      ["05","@m.poryvay",false],
      ["06","@g.kraynik",false],
      ["07","@r.gabibov",false],
      ["08","@m.poryvay",false],
    ];
    schedule.getRange(1,1,data.length,3).setValues(data);
    schedule.setColumnWidths(1,3,120);
    schedule.getRange("A1:C1").setFontWeight("bold").setBackground("#ffd966");
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
    queue.setColumnWidths(1,4,120);
    queue.getRange("A1:D1").setFontWeight("bold").setBackground("#c6e0b4");
  }

  let main = ss.getSheetByName("Sheet1");
  if (!main) {
    main = ss.insertSheet("Sheet1");
    main.getRange("A1:D1").setValues([["Name","Urgent","NonUrgent","DutyStatus"]]);
    main.getRange("A1:D1").setFontWeight("bold").setBackground("#d9e1f2");
  }

  // даём GSheets немного времени применить изменения
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
    if (!hour) continue;
    res[String(hour).padStart(2,"0")] = { name: name || "", manual: manual === true };
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

  let queue = getDutyQueue();
  if (!queue.length) {
    Logger.log("⚠️ Очередь пуста, остановка");
    return {};
  }

  const unavailable = getUnavailable();
  const current = getDutySchedule();

  // активные, не болеющие
  let pool = queue
    .filter(q => q.active && q.name && !unavailable.has(normalizeName(q.name)))
    .map(q => q.name);
  if (!pool.length) {
    pool = queue.filter(q => q.active && q.name).map(q => q.name);
  }
  if (!pool.length) return {};

  // считаем, сколько раз каждый уже дежурил (исторически)
  const pastStats = {};
  for (const h of Object.keys(current)) {
    const entry = current[h];
    if (entry && entry.name) {
      const n = normalizeName(entry.name);
      pastStats[n] = (pastStats[n] || 0) + 1;
    }
  }

  // изначально всем по 0
  const assignCounts = Object.fromEntries(pool.map(p => [p, pastStats[normalizeName(p)] || 0]));
  const result = {};
  let prev = null;

  for (const h of DUTY_HOURS) {
    const existing = current[h];
    if (existing && existing.manual && existing.name) {
      result[h] = { name: existing.name, manual: true };
      const n = normalizeName(existing.name);
      if (assignCounts.hasOwnProperty(n)) assignCounts[n]++;
      prev = existing.name;
      continue;
    }

    // кандидаты — все, кроме предыдущего
    const candidates = pool.filter(p => normalizeName(p) !== normalizeName(prev));

    // выбираем того, у кого МЕНЬШЕ всего дежурств
    candidates.sort((a, b) => {
      const diff = (assignCounts[normalizeName(a)] || 0) - (assignCounts[normalizeName(b)] || 0);
      if (diff !== 0) return diff;
      // при равенстве — случайный выбор (чтобы не циклилось по одному порядку)
      return Math.random() - 0.5;
    });

    const chosen = candidates[0] || pool[0];
    result[h] = { name: chosen, manual: false };
    assignCounts[normalizeName(chosen)] = (assignCounts[normalizeName(chosen)] || 0) + 1;
    prev = chosen;
  }

  writeDutySchedule(result);
  Logger.log("✅ Автопланирование завершено (равномерное распределение)");
  return result;
}


/* -------------------------
   3. Запись расписания
   ------------------------ */
function writeDutySchedule(schedule) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DutySchedule");
  if (!sheet) return;

  const out = [["Hour","Name","Manual"]];
  for (const h of DUTY_HOURS) {
    const item = schedule[h] || { name: "", manual: false };
    out.push([h, item.name, item.manual]);
  }

  sheet.clearContents();
  sheet.getRange(1,1,out.length,3).setValues(out);
  sheet.setColumnWidths(1,3,120);
  sheet.getRange("A1:C1").setFontWeight("bold").setBackground("#ffd966");
  SpreadsheetApp.flush();
}

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

  // текущий и следующий дежурный (нормализованные)
  const now = new Date();
  const curH = ("0" + now.getHours()).slice(-2);
  const nextH = ("0" + ((now.getHours() + 1) % 24)).slice(-2);
  const curM = now.getMinutes();

  const total = Math.max(sheet.getLastRow() - 1, 0);
  if (total === 0) return;

  const names = sheet.getRange(2,1,total,1).getValues().map(r => String(r[0] || "").trim());
  const statuses = sheet.getRange(2,4,total,1).getValues();

  const onDutyName = (duty[curH] && duty[curH].name) ? duty[curH].name : "";
  const comingName = (duty[nextH] && duty[nextH].name) ? duty[nextH].name : "";
  const onDutyNorm = normalizeName(onDutyName);
  const comingNorm = normalizeName(comingName);

  for (let i = 0; i < names.length; i++) {
    const nameNorm = normalizeName(names[i]);
    if (nameNorm && onDutyNorm && nameNorm === onDutyNorm) {
      statuses[i][0] = "on duty";
    } else if (nameNorm && comingNorm && nameNorm === comingNorm && curM >= 50) {
      statuses[i][0] = "duty coming";
    } else {
      // оставляем пустым — не трогаем другие ручные статусы, если тебе нужно удерживать вручную
      statuses[i][0] = "";
    }
  }

  sheet.getRange(2,4,total,1).setValues(statuses);
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

  // Принудительно применим последние правки перед чтением
  SpreadsheetApp.flush();
  Utilities.sleep(120);

  const unavailable = getUnavailable();

  if (queueSheet) {
    const rows = Math.max(queueSheet.getLastRow() - 1, 0);
    if (rows > 0) {
      const range = queueSheet.getRange(2,1,rows,4);
      const data = range.getValues();
      let changed = false;

      for (let i = 0; i < data.length; i++) {
        const rawName = String(data[i][1] || "").trim();
        if (!rawName) continue;

        const nNorm = normalizeName(rawName);
        const wasActive = data[i][3] === true;
        const nowActive = !unavailable.has(nNorm) && !unavailable.has("@" + nNorm);

        if (wasActive !== nowActive) {
          data[i][3] = nowActive;
          changed = true;
          const msg = nowActive ? `${rawName} снова активен` : `${rawName} временно исключён (sick/vacation)`;
          ss.toast(msg, "Duty Update", 3);
          Logger.log("🔔 " + msg);
        }
      }

      if (changed) {
        range.setValues(data);
        SpreadsheetApp.flush();
        Utilities.sleep(120);
      }
    }
  }

  // пересчитываем и применяем статусы
  autoDutyScheduler();
  Utilities.sleep(150);
  updateDutyStatus();

  ss.toast("✅ Очередь и дежурства обновлены", "Duty Scheduler", 3);
  Logger.log("✅ Полное обновление завершено");
}
