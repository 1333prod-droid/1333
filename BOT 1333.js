// ==================== BOT 1333 CAR CENTER v8.0 ====================
// ✅ v8 ЗМІНИ:
//   • recalcClientBalance() — авто-перерахунок балансу з абонементів
//   • handleMarkVisit — списує кошти з абонементу + оновлює баланс клієнта
//   • confirmSubscriptionRequest — оновлює баланс клієнта після активації
//   • sell_ callback — оновлює баланс клієнта після продажу
//   • updateClientBalance() — єдина точка оновлення балансу
//   • syncAllBalances() — ручна функція для синхронізації всіх балансів

const BOT_TOKEN = "8351312682:AAHKXJWOOcGuF1DTP9ulC_z_zTz8Jsq2mzk";
const ADMIN_ID = 240212361;
const SPREADSHEET_ID = "1Yut81wa4ZwXpCmL0K3wxhwjREP0TdWy5TVBB2GnAcLI";
const WEBAPP_URL = "https://script.google.com/macros/s/AKfycbw9mqrVSrAVeoMa4Np2SxYFRCxj5cW0uX1x7S50LddB7pXnz0PIvL7gPJHmLyIJ9986GQ/exec";
const PHONE = "+380 77 001 13 33";
const ADDRESS = "вул.Гвардійців-Широнінців 51/1";
const TZ = "Europe/Kyiv";
const scriptProperties = PropertiesService.getScriptProperties();

const SHEET_CLIENTS      = "👥 Клієнти";
const SHEET_TYPES        = "💎 Типи абонементів";
const SHEET_SUBS         = "📋 Абонементи клієнтів";
const SHEET_PRICE        = "💰 Прайс";
const SHEET_SCHEDULE     = "📅 Розклад";
const SHEET_BOOKINGS     = "📝 Записи на послуги";
const SHEET_HISTORY      = "📜 Історія";
const SHEET_SUB_REQUESTS = "📬 Заявки на абонементи";
const SHEET_SUB_REQUESTS_ALTS = ["🔔 Заявки абонементів","Заявки абонементів"];

const DAY_NAMES       = ["Неділя","Понеділок","Вівторок","Середа","Четвер","П'ятниця","Субота"];
const DAY_NAMES_SHORT = ["Нд","Пн","Вт","Ср","Чт","Пт","Сб"];

// ─────────────────────── HELPERS ───────────────────────────

function getSheet(name) {
  return SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(name);
}

function safeGetRangeValues(sheet, startRow, startCol, numRows, numCols) {
  if (!sheet) return [];
  const last = sheet.getLastRow();
  if (last < startRow) return [];
  const rows = Math.min(numRows, last - startRow + 1);
  if (rows <= 0) return [];
  return sheet.getRange(startRow, startCol, rows, numCols).getValues();
}

function fmtDate(d) {
  if (!d) return "";
  if (typeof d === 'string') return d;
  if (d instanceof Date) return Utilities.formatDate(d, TZ, "dd.MM.yyyy");
  return "";
}

function parseDate(value) {
  if (!value) return "";
  if (typeof value === 'string') return value;
  if (value instanceof Date) return Utilities.formatDate(value, TZ, "dd.MM.yyyy");
  if (typeof value === 'number') {
    try {
      const ms = (Math.floor(value) - 25569) * 86400 * 1000;
      const dt = new Date(ms);
      if (dt.getFullYear() > 1900 && dt.getFullYear() < 2100)
        return Utilities.formatDate(dt, TZ, "dd.MM.yyyy");
    } catch(e) {}
  }
  return "";
}

function fmt(num) { return Math.round(num * 100) / 100; }

// ──── NORMALIZE PHONE ────────────────────────────────────────
function normalizePhone(phone) {
  if (!phone) return "";
  let p = String(phone).replace(/[\s\-\(\)]/g, "");
  p = p.replace(/[^\d+]/g, "");
  if (p.startsWith("0") && p.length === 10) p = "+38" + p;
  if (p.startsWith("380") && !p.startsWith("+")) p = "+" + p;
  if (/^38\d{10}$/.test(p)) p = "+" + p;
  return p;
}

// ═══════════════════════════════════════════════════════════════
// ✅ КЛЮЧОВА ФУНКЦІЯ: Оновлення балансу клієнта в таблиці
// Баланс = сума залишків по ВСІХ активних абонементах клієнта
// ═══════════════════════════════════════════════════════════════
function updateClientBalance(clientId) {
  if (!clientId) return 0;
  
  const subsSheet = getSheet(SHEET_SUBS);
  const clientSheet = getSheet(SHEET_CLIENTS);
  if (!subsSheet || !clientSheet) return 0;
  
  // Рахуємо суму балансів по всіх активних абонементах
  const subsData = safeGetRangeValues(subsSheet, 2, 1, 500, 11);
  let totalBalance = 0;
  
  subsData.forEach(r => {
    if (!r[0]) return;
    if (String(r[0]).trim() !== String(clientId).trim()) return;
    const status = String(r[10]).trim();
    // Враховуємо тільки активні абонементи
    if (status.includes("Активний") || status.includes("✅")) {
      const bal = parseFloat(r[9]) || 0;
      totalBalance += bal;
    }
  });
  
  // Записуємо в колонку G (індекс 7) листа Клієнти
  const clientsData = safeGetRangeValues(clientSheet, 2, 1, 500, 1);
  for (let i = 0; i < clientsData.length; i++) {
    if (String(clientsData[i][0]).trim() === String(clientId).trim()) {
      clientSheet.getRange(i + 2, 7).setValue(totalBalance);
      Logger.log(`✅ Баланс оновлено: ${clientId} = ${totalBalance} грн`);
      return totalBalance;
    }
  }
  return totalBalance;
}

// ✅ Синхронізація балансів УСІХ клієнтів (запускати вручну або по тригеру)
function syncAllBalances() {
  const clientSheet = getSheet(SHEET_CLIENTS);
  if (!clientSheet) return;
  
  const clientsData = safeGetRangeValues(clientSheet, 2, 1, 500, 1);
  let count = 0;
  
  clientsData.forEach(r => {
    if (r[0]) {
      updateClientBalance(String(r[0]).trim());
      count++;
    }
  });
  
  Logger.log(`✅ Синхронізовано балансів: ${count} клієнтів`);
  return count;
}

// ✅ Списання з абонементу (при відвідуванні)
// Повертає true якщо успішно, false якщо недостатньо коштів
function deductFromSub(clientId, subId, cost) {
  if (!subId || !cost || cost <= 0) return true; // безкоштовно
  
  const sheet = getSheet(SHEET_SUBS);
  if (!sheet) return false;
  
  const data = safeGetRangeValues(sheet, 2, 1, 500, 11);
  
  for (let i = 0; i < data.length; i++) {
    if (!data[i][1]) continue;
    if (String(data[i][1]).trim() !== String(subId).trim()) continue;
    if (String(data[i][0]).trim() !== String(clientId).trim()) continue;
    
    const currentBalance = parseFloat(data[i][9]) || 0;
    const currentSpent   = parseFloat(data[i][8]) || 0;
    
    if (currentBalance < cost) {
      Logger.log(`❌ Недостатньо коштів: ${currentBalance} < ${cost}`);
      return false;
    }
    
    const newBalance = currentBalance - cost;
    const newSpent   = currentSpent + cost;
    
    sheet.getRange(i + 2, 9).setValue(newSpent);   // колонка I — витрачено
    sheet.getRange(i + 2, 10).setValue(newBalance); // колонка J — залишок
    
    // Якщо баланс 0 — закрити абонемент
    if (newBalance <= 0) {
      sheet.getRange(i + 2, 11).setValue("⛔ Вичерпано");
    }
    
    Logger.log(`✅ Списано ${cost} грн з ${subId}. Залишок: ${newBalance}`);
    
    // Оновлюємо загальний баланс клієнта
    updateClientBalance(clientId);
    return true;
  }
  
  Logger.log(`❌ Абонемент не знайдено: ${subId}`);
  return false;
}

// ──── AUTO-REGISTER CLIENT ────────────────────────────────────
function autoRegisterClient(name, phone) {
  if (!name || !phone) return null;
  const normPhone = normalizePhone(phone);
  
  const sheet = getSheet(SHEET_CLIENTS);
  if (!sheet) return null;
  
  const data = safeGetRangeValues(sheet, 2, 1, 500, 8);
  
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] && normalizePhone(String(data[i][2])) === normPhone) {
      Logger.log("✅ Клієнт вже є: " + data[i][0]);
      return String(data[i][0]);
    }
  }
  
  const row = getNextRow(SHEET_CLIENTS);
  const clientId = genClientId();
  
  try {
    sheet.getRange(row, 1).setValue(clientId);
    sheet.getRange(row, 2).setValue(name);
    sheet.getRange(row, 3).setValue(phone);
    sheet.getRange(row, 4).setValue("");
    sheet.getRange(row, 6).setValue(fmtDate(new Date()));
    sheet.getRange(row, 7).setValue(0);
    
    Logger.log("🆕 Авто-реєстрація: " + clientId + " " + name);
    
    sendMessage(ADMIN_ID,
      `👤 <b>НОВИЙ КЛІЄНТ (з сайту)</b>\n\n` +
      `🆔 ${clientId}\n👤 ${name}\n📱 ${phone}`
    );
    
    return clientId;
  } catch(e) {
    Logger.log("❌ autoRegisterClient error: " + e);
    return null;
  }
}

// ──── LINK TELEGRAM ID ────────────────────────────────────────
function linkTelegramId(phone, telegramId) {
  const normPhone = normalizePhone(phone);
  const sheet = getSheet(SHEET_CLIENTS);
  if (!sheet) return false;
  
  const data = safeGetRangeValues(sheet, 2, 1, 500, 8);
  
  for (let i = 0; i < data.length; i++) {
    if (!data[i][0]) continue;
    if (normalizePhone(String(data[i][2])) === normPhone) {
      const existingTg = String(data[i][3]).trim();
      if (!existingTg || existingTg === "" || existingTg === "0") {
        sheet.getRange(i + 2, 4).setValue(String(telegramId));
        Logger.log("✅ Прив'язано TG ID " + telegramId + " до " + data[i][1]);
        return true;
      } else {
        Logger.log("ℹ️ TG ID вже прив'язаний: " + existingTg);
        return false;
      }
    }
  }
  return false;
}

// ─────────────────────── ID GENERATORS ─────────────────────────

function genClientId() {
  const data = safeGetRangeValues(getSheet(SHEET_CLIENTS), 2, 1, 500, 1);
  let max = 0;
  data.forEach(r => { if (r[0]) { const m = String(r[0]).match(/C-(\d+)/); if (m) max = Math.max(max, +m[1]); }});
  return "C-" + String(max + 1).padStart(3, "0");
}
function genBookingId() {
  const data = safeGetRangeValues(getSheet(SHEET_BOOKINGS), 2, 1, 500, 1);
  let max = 0;
  data.forEach(r => { if (r[0]) { const m = String(r[0]).match(/B-(\d+)/); if (m) max = Math.max(max, +m[1]); }});
  return "B-" + String(max + 1).padStart(4, "0");
}
function genSubId() {
  const data = safeGetRangeValues(getSheet(SHEET_SUBS), 2, 1, 500, 2);
  let max = 0;
  data.forEach(r => { if (r[1]) { const m = String(r[1]).match(/A-(\d+)/); if (m) max = Math.max(max, +m[1]); }});
  return "A-" + String(max + 1).padStart(3, "0");
}
function genHistoryId() {
  const data = safeGetRangeValues(getSheet(SHEET_HISTORY), 2, 1, 500, 1);
  let max = 0;
  data.forEach(r => { if (r[0]) { const m = String(r[0]).match(/H-(\d+)/); if (m) max = Math.max(max, +m[1]); }});
  return "H-" + String(max + 1).padStart(4, "0");
}
function genSubRequestId() {
  const sheet = getSubRequestSheet();
  if (!sheet) return "SR-0001";
  const data = safeGetRangeValues(sheet, 2, 1, 500, 1);
  let max = 0;
  data.forEach(r => { if (r[0]) { const m = String(r[0]).match(/SR-(\d+)/); if (m) max = Math.max(max, +m[1]); }});
  return "SR-" + String(max + 1).padStart(4, "0");
}

function getNextRow(sheetName) {
  const sheet = getSheet(sheetName);
  if (!sheet) return 2;
  const data = safeGetRangeValues(sheet, 2, 1, 1000, 1);
  for (let i = 0; i < data.length; i++) { if (!data[i][0]) return i + 2; }
  return sheet.getLastRow() + 1;
}

// ─────────────────────── TELEGRAM API ──────────────────────────

function sendMessage(chatId, text, keyboard = null) {
  const url = `https://api.telegram.org/bot${BOT_TOKEN}/sendMessage`;
  const payload = { chat_id: chatId, text, parse_mode: "HTML" };
  if (keyboard) payload.reply_markup = JSON.stringify(keyboard);
  try {
    return JSON.parse(UrlFetchApp.fetch(url, {
      method: "post", contentType: "application/json",
      payload: JSON.stringify(payload), muteHttpExceptions: true
    }).getContentText());
  } catch(e) { Logger.log("sendMessage error: " + e); return null; }
}

function answerCallback(callbackId, text = "") {
  try {
    UrlFetchApp.fetch(`https://api.telegram.org/bot${BOT_TOKEN}/answerCallbackQuery`, {
      method: "post", contentType: "application/json",
      payload: JSON.stringify({ callback_query_id: callbackId, text }),
      muteHttpExceptions: true
    });
  } catch(e) {}
}

function createInlineKeyboard(rows) { return { inline_keyboard: rows }; }
function createReplyKeyboard(rows) { return { keyboard: rows, resize_keyboard: true, one_time_keyboard: false }; }

// ─────────────────────── CLIENT QUERIES ────────────────────────

function findClientByTelegram(telegramId) {
  const data = safeGetRangeValues(getSheet(SHEET_CLIENTS), 2, 1, 500, 8);
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] && String(data[i][3]).trim() === String(telegramId).trim()) {
      // ✅ Беремо актуальний баланс з абонементів, а не з кешу в таблиці
      const liveBalance = getClientLiveBalance(String(data[i][0]));
      return {
        clientId: data[i][0], name: data[i][1], phone: data[i][2],
        registered: parseDate(data[i][5]), balance: liveBalance
      };
    }
  }
  return null;
}

function findClientById(clientId) {
  const data = safeGetRangeValues(getSheet(SHEET_CLIENTS), 2, 1, 500, 8);
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(clientId).trim()) {
      const liveBalance = getClientLiveBalance(String(data[i][0]));
      return {
        clientId: data[i][0], name: data[i][1], phone: data[i][2],
        telegramId: String(data[i][3] || '').trim(),
        registered: parseDate(data[i][5]), balance: liveBalance
      };
    }
  }
  return null;
}

function findClientByPhone(phone) {
  const normPhone = normalizePhone(phone);
  const data = safeGetRangeValues(getSheet(SHEET_CLIENTS), 2, 1, 500, 8);
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] && normalizePhone(String(data[i][2])) === normPhone) {
      const liveBalance = getClientLiveBalance(String(data[i][0]));
      return {
        row: i + 2, clientId: data[i][0], name: data[i][1], phone: data[i][2],
        telegramId: String(data[i][3]).trim(), balance: liveBalance
      };
    }
  }
  return null;
}

// ✅ Живий баланс клієнта — рахується з абонементів в реальному часі
function getClientLiveBalance(clientId) {
  if (!clientId) return 0;
  const subsData = safeGetRangeValues(getSheet(SHEET_SUBS), 2, 1, 500, 11);
  let total = 0;
  subsData.forEach(r => {
    if (!r[0]) return;
    if (String(r[0]).trim() !== String(clientId).trim()) return;
    const status = String(r[10]).trim();
    if (status.includes("Активний") || status.includes("✅")) {
      total += parseFloat(r[9]) || 0;
    }
  });
  return total;
}

// ─────────────────────── DATA GETTERS ──────────────────────────

function getTypes() {
  const data = safeGetRangeValues(getSheet(SHEET_TYPES), 2, 1, 50, 6);
  return data.filter(r => r[0]).map(r => ({
    id: r[0], name: r[1], months: r[2], price: r[3], bonus: r[4], amount: r[5]
  }));
}

function getClientSubs(clientId) {
  const data = safeGetRangeValues(getSheet(SHEET_SUBS), 2, 1, 500, 11);
  return data
    .map((r, i) => ({ ...r, _row: i + 2 }))
    .filter(r => String(r[0]).trim() === String(clientId).trim() && r[1])
    .map(r => ({
      row: r._row, subId: r[1], typeName: r[4], amount: r[5],
      startDate: parseDate(r[6]), endDate: parseDate(r[7]),
      spent: parseFloat(r[8]) || 0, balance: parseFloat(r[9]) || 0,
      status: r[10]
    }));
}

function getClientHistory(clientId) {
  const data = safeGetRangeValues(getSheet(SHEET_HISTORY), 2, 1, 500, 8);
  return data.filter(r => String(r[2]).trim() === String(clientId).trim())
    .map(r => ({ date: parseDate(r[1]), service: r[5], cost: r[6] }));
}

function getClientBookings(clientId) {
  const sheet = getSheet(SHEET_BOOKINGS);
  if (!sheet) return [];
  const data = safeGetRangeValues(sheet, 2, 1, 500, 11);
  return data.filter(r => r[0] && String(r[5]).trim() === String(clientId).trim()).map(row => ({
    id: String(row[0]),
    date: row[1] instanceof Date ? fmtDate(row[1]) : String(row[1]).trim(),
    time: row[3] instanceof Date ? Utilities.formatDate(row[3], TZ, "HH:mm") : String(row[3]).trim(),
    box: String(row[4]), service: String(row[8]), status: String(row[9])
  }));
}

function getPriceByType(carType) {
  const data = safeGetRangeValues(getSheet(SHEET_PRICE), 2, 1, 100, 4);
  const services = {};
  data.forEach(r => {
    if (r[0] === carType && !services[r[1]])
      services[r[1]] = { name: r[1], duration: r[2], price: r[3] };
  });
  return services;
}

// ─────────────────────── SCHEDULE ──────────────────────────────

function getScheduleForDay(dayName) {
  const sheet = getSheet(SHEET_SCHEDULE);
  if (!sheet) return null;
  const data = safeGetRangeValues(sheet, 2, 1, 7, 8);
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === dayName) {
      return { day: data[i][0], box1_open: data[i][1], box1_close: data[i][2],
               box2_open: data[i][3], box2_close: data[i][4],
               box3_open: data[i][5], box3_close: data[i][6], status: data[i][7] };
    }
  }
  return null;
}

function parseTimeString(timeStr) {
  if (!timeStr || timeStr === "") return null;
  try {
    if (timeStr instanceof Date) {
      const s = Utilities.formatDate(timeStr, TZ, "HH:mm");
      const h = parseInt(s.split(":")[0]);
      return (isNaN(h) || h < 0 || h > 23) ? null : h;
    }
    const str = String(timeStr).trim();
    if (!str || str === "0") return null;
    const h = parseInt(str.split(str.includes(".") ? "." : ":")[0]);
    return (isNaN(h) || h < 0 || h > 23) ? null : h;
  } catch(e) { return null; }
}

function normalizeTime(timeStr) {
  if (!timeStr) return "";
  const str = String(timeStr).trim();
  if (str.includes(".")) return str.replace(".", ":").substring(0, 5);
  if (str.includes(":")) return str.substring(0, 5);
  return str.padStart(2, "0") + ":00";
}

function getDateSchedule(dateStr) {
  const [day, month, year] = dateStr.split(".");
  const d = new Date(parseInt(year), parseInt(month) - 1, parseInt(day));
  return getScheduleForDay(DAY_NAMES[d.getDay()]);
}

function getTimesForDate(dateStr) {
  const schedule = getDateSchedule(dateStr);
  if (!schedule) return [];
  const times = new Set();
  const addTimes = (open, close) => {
    const oh = parseTimeString(open), ch = parseTimeString(close);
    if (oh === null || ch === null) return;
    for (let h = oh; h < ch; h++) times.add(String(h).padStart(2, "0") + ":00");
  };
  addTimes(schedule.box1_open, schedule.box1_close);
  addTimes(schedule.box2_open, schedule.box2_close);
  addTimes(schedule.box3_open, schedule.box3_close);
  return Array.from(times).sort();
}

function getBoxesForTime(dateStr, timeStr) {
  const schedule = getDateSchedule(dateStr);
  if (!schedule) return [];
  const timeHour = parseTimeString(timeStr);
  if (timeHour === null) return [];
  const boxes = [];
  const checkBox = (open, close, name) => {
    const oh = parseTimeString(open), ch = parseTimeString(close);
    if (oh !== null && ch !== null && timeHour >= oh && timeHour < ch) boxes.push(name);
  };
  checkBox(schedule.box1_open, schedule.box1_close, "Box 1");
  checkBox(schedule.box2_open, schedule.box2_close, "Box 2");
  checkBox(schedule.box3_open, schedule.box3_close, "Box 3");
  return boxes;
}

// ─────────────────────── BOOKING ───────────────────────────────

function isBoxBooked(dateStr, timeStr, boxName) {
  const sheet = getSheet(SHEET_BOOKINGS);
  if (!sheet) return false;
  const data = safeGetRangeValues(sheet, 2, 1, 500, 11);
  const normInputTime = normalizeTime(timeStr);
  for (let i = 0; i < data.length; i++) {
    if (!data[i][0]) continue;
    let bDate = data[i][1];
    if (bDate instanceof Date) bDate = Utilities.formatDate(bDate, TZ, "dd.MM.yyyy");
    else bDate = String(bDate).trim();
    let bTime = data[i][3];
    if (bTime instanceof Date) bTime = Utilities.formatDate(bTime, TZ, "HH:mm");
    else bTime = String(bTime).trim();
    const bBox    = String(data[i][4]).trim();
    const bStatus = String(data[i][9]).trim();
    if (bDate === dateStr && bBox === boxName && bStatus.includes("Підтверджено")) {
      if (normalizeTime(bTime) === normInputTime) return true;
    }
  }
  return false;
}

function getAvailableBoxes(dateStr, timeStr) {
  return getBoxesForTime(dateStr, timeStr).filter(b => !isBoxBooked(dateStr, timeStr, b));
}

function isTimeFullyBooked(dateStr, timeStr) {
  const all = getBoxesForTime(dateStr, timeStr);
  return !all.length || all.every(b => isBoxBooked(dateStr, timeStr, b));
}

function saveBooking(dateStr, timeStr, boxName, clientId, clientName, phone, service) {
  const sheet = getSheet(SHEET_BOOKINGS);
  if (!sheet) return false;
  if (isBoxBooked(dateStr, timeStr, boxName)) {
    Logger.log("❌ Бокс вже зайнятий: " + boxName);
    return false;
  }
  const row = getNextRow(SHEET_BOOKINGS);
  const [day, month, year] = dateStr.split(".");
  const d = new Date(parseInt(year), parseInt(month) - 1, parseInt(day));
  const dayName = DAY_NAMES[d.getDay()];
  const normTime = normalizeTime(timeStr);
  try {
    sheet.getRange(row, 1).setValue(genBookingId());
    sheet.getRange(row, 2).setValue(dateStr);
    sheet.getRange(row, 3).setValue(dayName);
    sheet.getRange(row, 4).setValue(normTime);
    sheet.getRange(row, 5).setValue(boxName);
    sheet.getRange(row, 6).setValue(clientId);
    sheet.getRange(row, 7).setValue(clientName);
    sheet.getRange(row, 8).setValue(phone);
    sheet.getRange(row, 9).setValue(service);
    sheet.getRange(row, 10).setValue("✅ Підтверджено");
    sheet.getRange(row, 11).setValue("Автозапис");
    return true;
  } catch(e) { Logger.log("❌ saveBooking error: " + e); return false; }
}

// ─────────────────────── SUBSCRIPTION REQUESTS ─────────────────

function getSubRequestSheet() {
  let s = getSheet(SHEET_SUB_REQUESTS);
  if (s) return s;
  for (const alt of SHEET_SUB_REQUESTS_ALTS) { s = getSheet(alt); if (s) return s; }
  return null;
}

function saveSubscriptionRequest(clientId, clientName, telegramId, typeId, typeName, amount) {
  const sheet = getSubRequestSheet();
  if (!sheet) return false;
  const row = getNextRow(sheet.getName());
  const requestId = genSubRequestId();
  try {
    sheet.getRange(row, 1).setValue(requestId);
    sheet.getRange(row, 2).setValue(clientId);
    sheet.getRange(row, 3).setValue(clientName);
    sheet.getRange(row, 4).setValue(telegramId);
    sheet.getRange(row, 5).setValue(typeId);
    sheet.getRange(row, 6).setValue(typeName);
    sheet.getRange(row, 7).setValue(amount);
    sheet.getRange(row, 8).setValue(fmtDate(new Date()));
    sheet.getRange(row, 9).setValue("🆕 Нова заявка");
    return requestId;
  } catch(e) { Logger.log("❌ saveSubscriptionRequest: " + e); return false; }
}

function getSubscriptionRequests() {
  const sheet = getSubRequestSheet();
  if (!sheet) return [];
  return safeGetRangeValues(sheet, 2, 1, 500, 9).filter(r => r[0]).map((r, i) => ({
    row: i + 2, requestId: String(r[0]), clientId: String(r[1]), clientName: String(r[2]),
    telegramId: String(r[3]), typeId: String(r[4]), typeName: String(r[5]),
    amount: r[6] || 0, createdDate: parseDate(r[7]), status: String(r[8])
  }));
}

function updateSubscriptionRequestStatus(requestId, newStatus) {
  const sheet = getSubRequestSheet();
  if (!sheet) return false;
  const data = safeGetRangeValues(sheet, 2, 1, 500, 1);
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]) === String(requestId)) {
      sheet.getRange(i + 2, 9).setValue(newStatus);
      return true;
    }
  }
  return false;
}

function confirmSubscriptionRequest(requestId) {
  const req = getSubscriptionRequests().find(r => r.requestId === requestId);
  if (!req) return false;
  try {
    const sheet = getSheet(SHEET_SUBS);
    const row = getNextRow(SHEET_SUBS);
    const subId = genSubId();
    const pDate = new Date(), eDate = new Date();
    const type = getTypes().find(t => t.id === req.typeId);
    if (type) eDate.setMonth(eDate.getMonth() + (type.months || 1));
    
    sheet.getRange(row, 1).setValue(req.clientId);
    sheet.getRange(row, 2).setValue(subId);
    sheet.getRange(row, 3).setValue(req.typeId);
    sheet.getRange(row, 4).setValue(req.clientName);
    sheet.getRange(row, 5).setValue(req.typeName);
    sheet.getRange(row, 6).setValue(req.amount);
    sheet.getRange(row, 7).setValue(fmtDate(pDate));
    sheet.getRange(row, 8).setValue(fmtDate(eDate));
    sheet.getRange(row, 9).setValue(0);         // витрачено = 0
    sheet.getRange(row, 10).setValue(req.amount); // залишок = повна сума
    sheet.getRange(row, 11).setValue("✅ Активний");
    
    updateSubscriptionRequestStatus(requestId, "✅ Підтверджена");
    
    // ✅ ОНОВЛЮЄМО БАЛАНС КЛІЄНТА
    Utilities.sleep(300);
    updateClientBalance(req.clientId);
    
    return true;
  } catch(e) { Logger.log("❌ confirmSubscriptionRequest: " + e); return false; }
}

// ─────────────────────── STATE ─────────────────────────────────

function getState(userId) {
  try { const s = scriptProperties.getProperty('user_' + userId); return s ? JSON.parse(s) : null; }
  catch(e) { return null; }
}
function setState(userId, state) {
  try { scriptProperties.setProperty('user_' + userId, JSON.stringify(state)); } catch(e) {}
}
function deleteState(userId) { scriptProperties.deleteProperty('user_' + userId); }

// ─────────────────────── WEBHOOK ───────────────────────────────

function setWebhook() {
  const url = `https://api.telegram.org/bot${BOT_TOKEN}/setWebhook?url=${WEBAPP_URL}`;
  try { UrlFetchApp.fetch(url); Logger.log("✅ Webhook встановлено"); }
  catch(e) { Logger.log("❌ Webhook: " + e); }
}

function doPost(e) {
  try {
    const update = JSON.parse(e.postData.contents);
    if (update.message) handleMessage(update.message);
    if (update.callback_query) handleCallback(update.callback_query);
  } catch(err) { Logger.log("doPost error: " + err); }
}

// ─────────────────────── MESSAGE HANDLER ───────────────────────

function handleMessage(message) {
  const chatId = message.chat.id;
  const userId = message.from.id;
  const text = message.text || "";
  const firstName = message.from.first_name || "Користувач";

  if (message.contact) {
    const phone = message.contact.phone_number || "";
    const existingByPhone = findClientByPhone(phone);
    
    if (existingByPhone) {
      const linked = linkTelegramId(phone, userId);
      if (linked) Logger.log("🔗 Прив'язано TG ID до " + existingByPhone.clientId);
      const updatedClient = findClientByTelegram(userId);
      if (updatedClient) {
        sendMessage(chatId, `✅ <b>Ваш акаунт знайдено!</b>\n\n🆔 ${updatedClient.clientId}\n👤 ${updatedClient.name}\n💰 ${fmt(updatedClient.balance)} грн\n\nТепер ви можете користуватись ботом!`);
        showMainMenu(chatId, updatedClient);
        return;
      }
    }
    
    setState(userId, { action: "registration", step: 1, phone: phone });
    sendMessage(chatId, "📝 <b>Введіть ПІБ:</b>\n\n<i>Приклад: Іванов Іван Іванович</i>");
    return;
  }

  if (text === "/start") {
    deleteState(userId);
    sendWelcomeMessage(chatId, userId, firstName);
    return;
  }

  if (text === "/cancel") {
    deleteState(userId);
    const client = findClientByTelegram(userId);
    if (client) showMainMenu(chatId, client);
    else sendMessage(chatId, "❌ Скасовано");
    return;
  }

  if (userId === ADMIN_ID && text === "/admin") {
    showAdminMenu(chatId);
    return;
  }

  // ✅ Команда синхронізації балансів (тільки для адміна)
  if (userId === ADMIN_ID && text === "/syncbalances") {
    const count = syncAllBalances();
    sendMessage(chatId, `✅ Синхронізовано балансів: ${count} клієнтів`);
    return;
  }

  const client = findClientByTelegram(userId);
  const state = getState(userId);

  if (state && state.action === "registration" && state.step === 1) {
    const sheet = getSheet(SHEET_CLIENTS);
    const row = getNextRow(SHEET_CLIENTS);
    const clientId = genClientId();
    try {
      sheet.getRange(row, 1).setValue(clientId);
      sheet.getRange(row, 2).setValue(text);
      sheet.getRange(row, 3).setValue(state.phone);
      sheet.getRange(row, 4).setValue(String(userId));
      sheet.getRange(row, 6).setValue(fmtDate(new Date()));
      sheet.getRange(row, 7).setValue(0);
      Utilities.sleep(500);
      sendMessage(chatId, `✅ <b>Реєстрація готова!</b>\n\n🆔 ${clientId}\n👤 ${text}\n📱 ${state.phone}`);
      const newClient = findClientByTelegram(userId);
      if (newClient) showMainMenu(chatId, newClient);
    } catch(e) { sendMessage(chatId, "❌ Помилка: " + e.toString()); }
    deleteState(userId);
    return;
  }

  if (!client) {
    sendMessage(chatId, "❌ Ви не зареєстровані!\n\nВикористовуйте /start");
    return;
  }

  if (text === "📋 Мої абонементи") {
    const subs = getClientSubs(client.clientId);
    if (!subs.length) { sendMessage(chatId, "📋 <b>Абонементи не знайдені</b>"); return; }
    let msg = "📋 <b>АБОНЕМЕНТИ:</b>\n\n";
    subs.forEach((s, i) => {
      msg += `${i+1}. <b>${s.typeName}</b>\n` +
             `💰 Залишок: ${fmt(s.balance)} грн\n` +
             `✅ Витрачено: ${fmt(s.spent)} грн\n` +
             `📅 До: ${s.endDate}\n${s.status}\n\n`;
    });
    sendMessage(chatId, msg);
    return;
  }

  if (text === "💰 Прайс") {
    const keyboard = createInlineKeyboard([
      [{ text: "🚗 Sedan",    callback_data: "price_Sedan"   }],
      [{ text: "🚙 SUV",      callback_data: "price_SUV"     }],
      [{ text: "🚐 Мінівен",  callback_data: "price_Minivan" }]
    ]);
    sendMessage(chatId, "🚗 <b>Тип авто:</b>", keyboard);
    return;
  }

  if (text === "📅 Записатися") {
    const today = new Date();
    const dates = [];
    for (let i = 0; i < 7; i++) {
      const d = new Date(today);
      d.setDate(today.getDate() + i);
      const dateStr = Utilities.formatDate(d, TZ, "dd.MM.yyyy");
      const dayName = DAY_NAMES_SHORT[d.getDay()];
      dates.push({ text: `${dayName} ${dateStr}`, callback_data: `book_date_${dateStr}` });
    }
    const keyboard = createInlineKeyboard([
      [dates[0], dates[1]], [dates[2], dates[3]],
      [dates[4], dates[5]], [dates[6]]
    ]);
    sendMessage(chatId, "📅 <b>Оберіть дату:</b>", keyboard);
    return;
  }

  if (text === "👤 Мій профіль") {
    const subs = getClientSubs(client.clientId);
    const activeSubs = subs.filter(s => String(s.status).includes("Активний") || String(s.status).includes("✅"));
    let msg = `👤 <b>ПРОФІЛЬ</b>\n\n🆔 ${client.clientId}\n👤 ${client.name}\n📱 ${client.phone}\n💰 Баланс: <b>${fmt(client.balance)} грн</b>`;
    if (activeSubs.length) {
      msg += `\n\n📋 <b>Активні абонементи:</b>`;
      activeSubs.forEach(s => { msg += `\n• ${s.typeName}: ${fmt(s.balance)} грн`; });
    }
    sendMessage(chatId, msg);
    return;
  }

  if (text === "📞 Контакти") {
    sendMessage(chatId, `📞 <b>КОНТАКТИ</b>\n\n📱 ${PHONE}\n📍 ${ADDRESS}`);
    return;
  }

  if (text === "💎 Абонементи") {
    const types = getTypes();
    let msg = "💎 <b>АБОНЕМЕНТИ:</b>\n\n";
    const buttons = [];
    types.forEach(t => {
      msg += `<b>${t.name}</b>\n💰 ${t.price} грн | ${t.amount} грн на рахунок\n⏱ ${t.months} міс\n\n`;
      buttons.push([{ text: `Замовити ${t.name}`, callback_data: `order_${t.id}` }]);
    });
    sendMessage(chatId, msg, createInlineKeyboard(buttons));
    return;
  }

  if (text === "📊 Статистика") {
    const subs = getClientSubs(client.clientId);
    const history = getClientHistory(client.clientId);
    let totalSpent = 0;
    history.forEach(h => { totalSpent += parseFloat(h.cost) || 0; });
    const activeSubs = subs.filter(s => String(s.status).includes("Активний") || String(s.status).includes("✅"));
    sendMessage(chatId,
      `📊 <b>СТАТИСТИКА</b>\n\n` +
      `📋 Абонементів: ${subs.length}\n` +
      `✅ Активних: ${activeSubs.length}\n` +
      `💰 Поточний баланс: ${fmt(client.balance)} грн\n` +
      `✅ Візитів: ${history.length}\n` +
      `💸 Витрачено всього: ${fmt(totalSpent)} грн`
    );
    return;
  }

  if (userId === ADMIN_ID) {
    if (text === "➕ Додати клієнта")          { setState(userId, { action: "add_client", step: 1 }); sendMessage(chatId, "➕ <b>ПІБ:</b>"); return; }
    if (text === "🎫 Продати абонемент")         { setState(userId, { action: "sell_sub", step: 1 }); sendMessage(chatId, "🎫 <b>ID клієнта:</b>"); return; }
    if (text === "✅ Відмітити відвідування")    { setState(userId, { action: "mark_visit", step: 1 }); sendMessage(chatId, "✅ <b>ID клієнта:</b>"); return; }
    if (text === "👥 Список клієнтів")           { showClientsList(chatId); return; }
    if (text === "📊 Аналітика")                 { showAnalytics(chatId); return; }
  }

  if (state) {
    if (userId === ADMIN_ID) {
      if (state.action === "add_client")  handleAddClient(chatId, userId, text, state);
      if (state.action === "sell_sub")    handleSellSub(chatId, userId, text, state);
      if (state.action === "mark_visit")  handleMarkVisit(chatId, userId, text, state);
    }
    return;
  }

  sendMessage(chatId, "Використовуй меню 👇");
}

// ─────────────────────── CALLBACK HANDLER ──────────────────────

function handleCallback(callbackQuery) {
  const chatId = callbackQuery.message.chat.id;
  const userId = callbackQuery.from.id;
  const data   = callbackQuery.data;

  answerCallback(callbackQuery.id);

  const client = findClientByTelegram(userId);
  const state  = getState(userId);

  if (data.startsWith("price_")) {
    const services = getPriceByType(data.replace("price_", ""));
    if (!Object.keys(services).length) { sendMessage(chatId, "❌ Послуги не знайдені"); return; }
    let msg = `🚗 <b>${data.replace("price_","")}</b>\n\n`;
    Object.values(services).forEach(s => { msg += `• <b>${s.name}</b>\n  ⏱ ${s.duration} хв | 💰 ${s.price} грн\n\n`; });
    sendMessage(chatId, msg);
    return;
  }

  if (data.startsWith("book_date_")) {
    const dateStr = data.replace("book_date_", "");
    const times = getTimesForDate(dateStr);
    if (!times.length) { sendMessage(chatId, "❌ На цю дату нема часу. Оберіть іншу."); return; }
    setState(userId, { action: "booking", date: dateStr });
    const buttons = times.map(t => [{ text: `⏰ ${t}`, callback_data: `book_time_${t}` }]);
    sendMessage(chatId, `📅 ${dateStr}\n\n⏰ <b>Оберіть час:</b>`, createInlineKeyboard(buttons));
    return;
  }

  if (data.startsWith("book_time_")) {
    const timeStr = data.replace("book_time_", "");
    if (!state?.date) { sendMessage(chatId, "❌ Помилка. Почніть знову."); return; }
    const boxes = getAvailableBoxes(state.date, timeStr);
    if (!boxes.length) { sendMessage(chatId, "❌ Нема вільних боксів. Оберіть інший час."); return; }
    setState(userId, { action: "booking", date: state.date, time: timeStr });
    const buttons = boxes.map(b => [{ text: b, callback_data: `book_box_${b}` }]);
    sendMessage(chatId, `⏰ ${timeStr}\n\n🏠 <b>Бокс:</b>`, createInlineKeyboard(buttons));
    return;
  }

  if (data.startsWith("book_box_")) {
    const box = data.replace("book_box_", "");
    if (!state || !client) { sendMessage(chatId, "❌ Помилка"); return; }
    const ok = saveBooking(state.date, state.time, box, client.clientId, client.name, client.phone, "Запис через бот");
    if (ok) {
      sendMessage(chatId, `✅ <b>Запис готовий!</b>\n\n📅 ${state.date}\n⏰ ${state.time}\n🏠 ${box}\n\nОчікуємо на вас! 🚗`);
      sendMessage(ADMIN_ID, `📝 <b>НОВИЙ ЗАПИС!</b>\n\n👤 ${client.name}\n📅 ${state.date} ${state.time}\n🏠 ${box}`);
    } else {
      sendMessage(chatId, "❌ Бокс вже зайнятий. Спробуйте інший час.");
      return;
    }
    deleteState(userId);
    Utilities.sleep(300);
    showMainMenu(chatId, client);
    return;
  }

  if (data.startsWith("order_")) {
    const typeId = data.replace("order_", "");
    if (!client) { sendMessage(chatId, "❌ Помилка"); return; }
    const type = getTypes().find(t => t.id === typeId);
    const requestId = saveSubscriptionRequest(client.clientId, client.name, String(userId), typeId, type?.name || typeId, type?.amount || 0);
    if (requestId) {
      sendMessage(ADMIN_ID,
        `🔔 <b>НОВА ЗАЯВКА НА АБОНЕМЕНТ!</b>\n\n👤 ${client.name}\n🆔 ${client.clientId}\n📱 ${client.phone}\n💎 <b>${type?.name || typeId}</b>\n💰 ${type?.amount || 0} грн\n📅 ID: ${requestId}`,
        createInlineKeyboard([[
          { text: "✅ Підтвердити", callback_data: `confirm_sub_${requestId}` },
          { text: "❌ Відхилити",   callback_data: `reject_sub_${requestId}`  }
        ]])
      );
      sendMessage(chatId, `✅ Заявка відправлена!\n\n💎 ${type?.name || typeId}\n⏳ Чекаємо підтвердження...`);
    } else {
      sendMessage(chatId, "❌ Помилка при створенні заявки");
    }
    return;
  }

  if (data.startsWith("confirm_sub_")) {
    const requestId = data.replace("confirm_sub_", "");
    const ok = confirmSubscriptionRequest(requestId);
    answerCallback(callbackQuery.id, ok ? "✅ Підтверджено!" : "❌ Помилка");
    sendMessage(ADMIN_ID, ok ? `✅ Заявка ${requestId} підтверджена!` : `❌ Помилка підтвердження ${requestId}`);
    if (ok) {
      const req = getSubscriptionRequests().find(r => r.requestId === requestId);
      if (req?.telegramId && req.telegramId !== 'undefined')
        sendMessage(req.telegramId, `🎉 <b>Абонемент активовано!</b>\n\n💎 ${req.typeName}\n💰 ${req.amount} грн`);
    }
    return;
  }

  if (data.startsWith("reject_sub_")) {
    const requestId = data.replace("reject_sub_", "");
    const ok = updateSubscriptionRequestStatus(requestId, "❌ Відхилена");
    answerCallback(callbackQuery.id, "❌ Відхилено");
    sendMessage(ADMIN_ID, `❌ Заявка ${requestId} відхилена`);
    if (ok) {
      const req = getSubscriptionRequests().find(r => r.requestId === requestId);
      if (req?.telegramId && req.telegramId !== 'undefined')
        sendMessage(req.telegramId, `❌ <b>Заявку відхилено</b>\n\n${req.typeName}\n\n💬 Зв'яжіться: ${PHONE}`);
    }
    return;
  }

  // ✅ ВИПРАВЛЕНО: продаж абонементу + оновлення балансу
  if (data.startsWith("sell_")) {
    const parts = data.replace("sell_", "").split("_");
    const clientId = parts[0];
    const typeId = parts.slice(1).join("_");
    const tClient = findClientById(clientId);
    const type = getTypes().find(t => String(t.id) === String(typeId));
    if (!type || !tClient) { sendMessage(chatId, "❌ Помилка: клієнт або тип не знайдено"); return; }
    try {
      const sheet = getSheet(SHEET_SUBS);
      const row = getNextRow(SHEET_SUBS);
      const subId = genSubId();
      const eDate = new Date();
      eDate.setMonth(eDate.getMonth() + (type.months || 1));
      sheet.getRange(row, 1).setValue(clientId);
      sheet.getRange(row, 2).setValue(subId);
      sheet.getRange(row, 3).setValue(typeId);
      sheet.getRange(row, 4).setValue(tClient.name);
      sheet.getRange(row, 5).setValue(type.name);
      sheet.getRange(row, 6).setValue(type.amount || 0);
      sheet.getRange(row, 7).setValue(fmtDate(new Date()));
      sheet.getRange(row, 8).setValue(fmtDate(eDate));
      sheet.getRange(row, 9).setValue(0);
      sheet.getRange(row, 10).setValue(type.amount || 0);
      sheet.getRange(row, 11).setValue("✅ Активний");
      
      // ✅ ОНОВЛЮЄМО БАЛАНС КЛІЄНТА
      Utilities.sleep(300);
      const newBalance = updateClientBalance(clientId);
      
      sendMessage(chatId, `✅ <b>Продано!</b>\n💎 ${type.name}\n💰 ${type.amount} грн\n👤 ${tClient.name}\n💳 Новий баланс: ${fmt(newBalance)} грн`);
      
      if (tClient.telegramId && tClient.telegramId !== '0' && tClient.telegramId !== '')
        sendMessage(tClient.telegramId, `🎉 <b>Ви придбали ${type.name}!</b>\n\n💰 ${type.amount} грн на рахунку\n📅 До: ${fmtDate(eDate)}\n\nДякуємо! 🚗`);
    } catch(e) { sendMessage(chatId, "❌ Помилка: " + e); }
    return;
  }

  if (data.startsWith("visit_sub_")) {
    const [clientId, subId] = data.replace("visit_sub_", "").split("|");
    const curState = getState(userId);
    if (curState?.action === "mark_visit") {
      setState(userId, { ...curState, step: 2, clientId, subId });
      sendMessage(chatId, `🎫 <b>${subId}</b>\n\n🔧 <b>Послуга:</b>`);
    }
    return;
  }
}

// ─────────────────────── ADMIN HANDLERS ────────────────────────

function handleAddClient(chatId, userId, text, state) {
  if (state.step === 1) {
    setState(userId, { action: "add_client", step: 2, name: text });
    sendMessage(chatId, "📱 <b>Телефон:</b>");
  } else if (state.step === 2) {
    setState(userId, { action: "add_client", step: 3, name: state.name, phone: text });
    sendMessage(chatId, "💬 <b>Telegram ID (або '-'):</b>");
  } else if (state.step === 3) {
    try {
      const sheet = getSheet(SHEET_CLIENTS);
      const row = getNextRow(SHEET_CLIENTS);
      const clientId = genClientId();
      sheet.getRange(row, 1).setValue(clientId);
      sheet.getRange(row, 2).setValue(state.name);
      sheet.getRange(row, 3).setValue(state.phone);
      sheet.getRange(row, 4).setValue(text === "-" ? "" : text);
      sheet.getRange(row, 6).setValue(fmtDate(new Date()));
      sheet.getRange(row, 7).setValue(0);
      deleteState(userId);
      sendMessage(chatId, `✅ ${clientId} — ${state.name} — додано!`);
      showAdminMenu(chatId);
    } catch(e) { sendMessage(chatId, "❌ Помилка: " + e); }
  }
}

function handleSellSub(chatId, userId, text, state) {
  if (state.step === 1) {
    const c = findClientById(text.trim());
    if (!c) { sendMessage(chatId, `❌ Клієнт ${text} не знайдений`); return; }
    setState(userId, { action: "sell_sub", step: 2, clientId: text.trim() });
    const types = getTypes();
    if (!types.length) { sendMessage(chatId, "❌ Типи абонементів не знайдено"); return; }
    const buttons = types.map(t => [{ text: `💎 ${t.name} — ${t.amount} грн`, callback_data: `sell_${text.trim()}_${t.id}` }]);
    sendMessage(chatId, `👤 ${c.name}\n💰 Поточний баланс: ${fmt(c.balance)} грн\n\n💎 Оберіть тип:`, createInlineKeyboard(buttons));
  }
}

// ✅ ВИПРАВЛЕНО: handleMarkVisit тепер списує кошти з абонементу
function handleMarkVisit(chatId, userId, text, state) {
  if (state.step === 1) {
    const c = findClientById(text.trim());
    if (!c) { sendMessage(chatId, "❌ Клієнт не знайдений"); return; }
    const subs = getClientSubs(text.trim()).filter(s =>
      String(s.status).includes("Активний") || String(s.status).includes("✅")
    );
    if (!subs.length) {
      sendMessage(chatId, `❌ У ${c.name} нема активних абонементів!\n💰 Баланс: ${fmt(c.balance)} грн`);
      deleteState(userId);
      return;
    }
    setState(userId, { action: "mark_visit", step: 2, clientId: text.trim(), clientName: c.name });
    const buttons = subs.map((s, i) => [{
      text: `${i+1}. ${s.typeName} | 💰 ${fmt(s.balance)} грн`,
      callback_data: `visit_sub_${text.trim()}|${s.subId}`
    }]);
    sendMessage(chatId, `👤 <b>${c.name}</b>\n💰 Загальний баланс: ${fmt(c.balance)} грн\n\n🎫 <b>Оберіть абонемент:</b>`, createInlineKeyboard(buttons));
  }
  else if (state.step === 2) {
    setState(userId, { ...state, step: 3, service: text });
    sendMessage(chatId, "💰 <b>Сума списання (грн):</b>");
  }
  else if (state.step === 3) {
    const cost = parseFloat(text) || 0;
    if (cost <= 0) { sendMessage(chatId, "❌ Введіть суму більше 0"); return; }
    setState(userId, { ...state, step: 4, cost });
    sendMessage(chatId, `📝 <b>Коментар</b> (або '-'):`);
  }
  else if (state.step === 4) {
    try {
      // ✅ Списуємо з абонементу
      const deducted = deductFromSub(state.clientId, state.subId, state.cost);
      if (!deducted) {
        sendMessage(chatId, `❌ Недостатньо коштів на абонементі!\nПеревірте баланс та повторіть.`);
        deleteState(userId);
        return;
      }
      
      // Записуємо в історію
      const sheet = getSheet(SHEET_HISTORY);
      const row = getNextRow(SHEET_HISTORY);
      sheet.getRange(row, 1).setValue(genHistoryId());
      sheet.getRange(row, 2).setValue(fmtDate(new Date()));
      sheet.getRange(row, 3).setValue(state.clientId);
      sheet.getRange(row, 4).setValue(state.clientName);
      sheet.getRange(row, 5).setValue(state.subId || "");
      sheet.getRange(row, 6).setValue(state.service);
      sheet.getRange(row, 7).setValue(state.cost);
      sheet.getRange(row, 8).setValue(text === "-" ? "" : text);
      
      // Отримуємо актуальний баланс після списання
      const newBalance = getClientLiveBalance(state.clientId);
      
      deleteState(userId);
      sendMessage(chatId,
        `✅ <b>Відвідування записано!</b>\n\n` +
        `👤 ${state.clientName}\n` +
        `🔧 ${state.service}\n` +
        `💸 Списано: ${state.cost} грн\n` +
        `💰 Залишок на рахунку: ${fmt(newBalance)} грн`
      );
      
      // Повідомляємо клієнта якщо є TG
      const clientData = findClientById(state.clientId);
      if (clientData?.telegramId && clientData.telegramId !== '0' && clientData.telegramId !== '') {
        sendMessage(clientData.telegramId,
          `🚗 <b>Відвідування зафіксовано!</b>\n\n` +
          `🔧 ${state.service}\n` +
          `💸 Списано: ${state.cost} грн\n` +
          `💰 Залишок: ${fmt(newBalance)} грн`
        );
      }
      
      showAdminMenu(chatId);
    } catch(e) { sendMessage(chatId, "❌ Помилка: " + e); }
  }
}

// ─────────────────────── ADMIN VIEWS ───────────────────────────

function showClientsList(chatId) {
  const clientSheet = getSheet(SHEET_CLIENTS);
  const data = safeGetRangeValues(clientSheet, 2, 1, 100, 8);
  let msg = "👥 <b>КЛІЄНТИ:</b>\n\n", count = 0;
  data.forEach(r => {
    if (r[0]) {
      count++;
      const liveBalance = getClientLiveBalance(String(r[0]));
      msg += `${count}. <b>${r[1]}</b> (${r[0]})\n📱 ${r[2]}\n💰 ${fmt(liveBalance)} грн\n\n`;
    }
  });
  if (!count) msg = "❌ Клієнтів не знайдено";
  sendMessage(chatId, msg);
}

function showAnalytics(chatId) {
  const sData = getSheet(SHEET_SUBS)    ? safeGetRangeValues(getSheet(SHEET_SUBS),    2, 1, 500, 11) : [];
  const hData = getSheet(SHEET_HISTORY) ? safeGetRangeValues(getSheet(SHEET_HISTORY), 2, 1, 500, 8)  : [];
  const cData = getSheet(SHEET_CLIENTS) ? safeGetRangeValues(getSheet(SHEET_CLIENTS), 2, 1, 500, 8)  : [];
  
  let totalSold = 0, totalSpent = 0, totalBalance = 0;
  sData.forEach(r => {
    if (!r[0]) return;
    totalSold    += parseFloat(r[5]) || 0;
    totalSpent   += parseFloat(r[8]) || 0;
    const status = String(r[10]).trim();
    if (status.includes("Активний") || status.includes("✅")) {
      totalBalance += parseFloat(r[9]) || 0;
    }
  });
  
  let visitTotal = 0;
  hData.forEach(r => { visitTotal += parseFloat(r[6]) || 0; });
  
  const activeClients = cData.filter(r => r[0]).length;
  const activeSubs = sData.filter(r => r[0] && (String(r[10]).includes("Активний") || String(r[10]).includes("✅"))).length;
  
  sendMessage(chatId,
    `📊 <b>АНАЛІТИКА</b>\n\n` +
    `💰 Продано абонементів: ${fmt(totalSold)} грн\n` +
    `✅ Відпрацьовано: ${fmt(totalSpent)} грн\n` +
    `💳 Залишок у клієнтів: ${fmt(totalBalance)} грн\n\n` +
    `👥 Клієнтів: ${activeClients}\n` +
    `📋 Активних абонементів: ${activeSubs}\n` +
    `✅ Візитів всього: ${hData.filter(r => r[0]).length}`
  );
}

function sendWelcomeMessage(chatId, userId, firstName) {
  const client = findClientByTelegram(userId);
  if (client) { showMainMenu(chatId, client); return; }
  if (userId === ADMIN_ID) { showAdminMenu(chatId); return; }
  const keyboard = createReplyKeyboard([[{ text: "📱 Поділитись номером", request_contact: true }]]);
  sendMessage(chatId,
    `🚗 <b>1333 Car Center</b>\n\n` +
    `Якщо ви реєструвались через сайт — натисніть кнопку нижче, і ваш Telegram буде автоматично прив'язаний!\n\n` +
    `Натисніть 👇 щоб продовжити:`,
    keyboard
  );
}

function showMainMenu(chatId, client) {
  const keyboard = createReplyKeyboard([
    [{ text: "📋 Мої абонементи" }, { text: "💎 Абонементи" }],
    [{ text: "💰 Прайс" },          { text: "📅 Записатися" }],
    [{ text: "👤 Мій профіль" },    { text: "📞 Контакти"   }],
    [{ text: "📊 Статистика" }]
  ]);
  sendMessage(chatId, `👋 Привіт, ${client.name}!\n💰 Баланс: <b>${fmt(client.balance)} грн</b>`, keyboard);
}

function showAdminMenu(chatId) {
  const keyboard = createReplyKeyboard([
    [{ text: "➕ Додати клієнта"         }, { text: "🎫 Продати абонемент"   }],
    [{ text: "✅ Відмітити відвідування" }, { text: "👥 Список клієнтів"     }],
    [{ text: "📊 Аналітика" }]
  ]);
  sendMessage(chatId, "👨‍💼 <b>АДМІН ПАНЕЛЬ</b>\n\n/syncbalances — синхронізувати всі баланси", keyboard);
}

// ─────────────────────── doGet (REST API) ──────────────────────

function doGet(e) {
  const action = e.parameter.action || '';
  let result = {};

  if (action === 'getAvailableTimesForDate') {
    const dateStr = e.parameter.date;
    const allTimes = getTimesForDate(dateStr);
    result = {
      times: allTimes.map(time => {
        const boxes     = getBoxesForTime(dateStr, time);
        const freeBoxes = boxes.filter(b => !isBoxBooked(dateStr, time, b));
        return { time, available: freeBoxes.length > 0, freeBoxes };
      }),
      date: dateStr
    };
  }
  else if (action === 'getBoxesForDateTime') {
    result = { boxes: getAvailableBoxes(e.parameter.date, e.parameter.time) };
  }
  else if (action === 'getAllBookings') {
    const sheet = getSheet(SHEET_BOOKINGS);
    const data  = safeGetRangeValues(sheet, 2, 1, 500, 11);
    result = {
      bookings: data.filter(r => r[0]).map((row, idx) => ({
        row: idx + 2, id: String(row[0]),
        date: row[1] instanceof Date ? fmtDate(row[1]) : String(row[1]).trim(),
        day: String(row[2]),
        time: row[3] instanceof Date ? Utilities.formatDate(row[3], TZ, "HH:mm") : String(row[3]).trim(),
        box: String(row[4]), clientId: String(row[5]), clientName: String(row[6]),
        phone: String(row[7]), service: String(row[8]), status: String(row[9]),
        note: String(row[10] || '')
      }))
    };
  }
  else if (action === 'getAllClients') {
    const sheet = getSheet(SHEET_CLIENTS);
    const data  = safeGetRangeValues(sheet, 2, 1, 500, 8);
    result = {
      clients: data.filter(r => r[0]).map(row => ({
        clientId:   String(row[0]),
        name:       String(row[1]),
        phone:      String(row[2]),
        telegramId: String(row[3] || ''),
        registered: parseDate(row[5]),
        balance:    getClientLiveBalance(String(row[0]))  // ✅ живий баланс
      }))
    };
  }
  else if (action === 'getClientFull') {
    const cid = e.parameter.clientId;
    result = {
      client:   findClientById(cid),
      subs:     getClientSubs(cid),
      history:  getClientHistory(cid),
      bookings: getClientBookings(cid)
    };
  }
  else if (action === 'getAllSubs') {
    const sheet = getSheet(SHEET_SUBS);
    const data  = safeGetRangeValues(sheet, 2, 1, 500, 11);
    result = {
      subs: data.filter(r => r[0]).map((row, idx) => ({
        row: idx + 2, clientId: String(row[0]), subId: String(row[1]),
        typeId: String(row[2]), clientName: String(row[3]), typeName: String(row[4]),
        amount: row[5] || 0, startDate: parseDate(row[6]), endDate: parseDate(row[7]),
        spent: row[8] || 0, balance: row[9] || 0, status: String(row[10])
      }))
    };
  }
  else if (action === 'updateBookingStatus') {
    const sheet = getSheet(SHEET_BOOKINGS);
    const data  = safeGetRangeValues(sheet, 2, 1, 500, 1);
    let updated = false;
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]) === e.parameter.bookingId) {
        sheet.getRange(i + 2, 10).setValue(e.parameter.status);
        updated = true; break;
      }
    }
    result = { success: updated };
  }
  else if (action === 'updateSubStatus') {
    const sheet = getSheet(SHEET_SUBS);
    const data  = safeGetRangeValues(sheet, 2, 1, 500, 2);
    let updated = false;
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][1]) === e.parameter.subId) {
        sheet.getRange(i + 2, 11).setValue(e.parameter.status);
        updated = true; break;
      }
    }
    result = { success: updated };
  }
  else if (action === 'saveBooking') {
    const name  = e.parameter.name  || 'Гість';
    const phone = e.parameter.phone || '';
    const clientId = autoRegisterClient(name, phone) || ('web-' + Date.now());
    const saved = saveBooking(
      e.parameter.date, e.parameter.time, e.parameter.box,
      clientId, name, phone,
      (e.parameter.service || '') + (e.parameter.car ? ' | ' + e.parameter.car : '')
    );
    if (saved) {
      sendMessage(ADMIN_ID,
        `🌐 <b>ЗАПИС З САЙТУ!</b>\n\n` +
        `👤 ${name}\n📱 ${phone}\n` +
        `📅 ${e.parameter.date} ${e.parameter.time}\n` +
        `🏠 ${e.parameter.box}\n🔧 ${e.parameter.service}\n` +
        `🚗 ${e.parameter.car||'—'} · ${e.parameter.carType||'—'}`
      );
    }
    result = { success: saved };
  }
  else if (action === 'saveSubOrder') {
    try {
      const name  = e.parameter.name  || 'Гість';
      const phone = e.parameter.phone || '';
      const clientId = autoRegisterClient(name, phone) || ('web-' + Date.now());
      const requestId = saveSubscriptionRequest(
        clientId, name, e.parameter.tg || '',
        e.parameter.subId, e.parameter.subName,
        Number(e.parameter.price) || 0
      );
      if (requestId) {
        sendMessage(ADMIN_ID,
          `💎 <b>НОВА ЗАЯВКА НА АБОНЕМЕНТ!</b>\n\n` +
          `👤 ${name}\n📱 ${phone}\n💬 Telegram: ${e.parameter.tg||'—'}\n\n` +
          `🏷 <b>${e.parameter.subName}</b>\n📅 ${e.parameter.months} міс.\n` +
          `💰 ${Number(e.parameter.price).toLocaleString('uk-UA')} грн → ${Number(e.parameter.account).toLocaleString('uk-UA')} грн (+${e.parameter.bonus})\n\n` +
          `ID заявки: ${requestId}`,
          createInlineKeyboard([[
            { text: "✅ Підтвердити", callback_data: `confirm_sub_${requestId}` },
            { text: "❌ Відхилити",   callback_data: `reject_sub_${requestId}`  }
          ]])
        );
        result = { success: true, requestId };
      } else {
        result = { success: false, error: 'Не вдалося зберегти заявку' };
      }
    } catch(err) { result = { success: false, error: err.toString() }; }
  }
  else if (action === 'getAllSubRequests') {
    result = { requests: getSubscriptionRequests() };
  }
  else if (action === 'updateSubRequestStatus') {
    result = { success: updateSubscriptionRequestStatus(e.parameter.requestId, e.parameter.status) };
  }
  else if (action === 'confirmSubRequest') {
    const ok = confirmSubscriptionRequest(e.parameter.requestId);
    if (ok) {
      const req = getSubscriptionRequests().find(r => r.requestId === e.parameter.requestId);
      if (req?.telegramId && req.telegramId !== 'undefined')
        sendMessage(req.telegramId, `🎉 <b>Заявка підтверджена!</b>\n\n💎 ${req.typeName}\n💰 ${req.amount} грн`);
    }
    result = { success: ok };
  }
  else if (action === 'rejectSubRequest') {
    const ok = updateSubscriptionRequestStatus(e.parameter.requestId, '❌ Відхилена');
    if (ok) {
      const req = getSubscriptionRequests().find(r => r.requestId === e.parameter.requestId);
      if (req?.telegramId && req.telegramId !== 'undefined')
        sendMessage(req.telegramId, `❌ <b>Заявку відхилено</b>\n\n${req.typeName}\n\n💬 ${PHONE}`);
    }
    result = { success: ok };
  }
  else if (action === 'autoRegister') {
    const name  = e.parameter.name  || '';
    const phone = e.parameter.phone || '';
    if (!name || !phone) {
      result = { success: false, error: 'name and phone required' };
    } else {
      const clientId = autoRegisterClient(name, phone);
      result = { success: !!clientId, clientId: clientId || null };
    }
  }
  // ✅ НОВИЙ endpoint: синхронізація балансів
  else if (action === 'syncBalances') {
    const count = syncAllBalances();
    result = { success: true, synced: count };
  }
  // ✅ НОВИЙ endpoint: отримати баланс клієнта
  else if (action === 'getClientBalance') {
    const cid = e.parameter.clientId;
    const balance = getClientLiveBalance(cid);
    result = { clientId: cid, balance };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}
