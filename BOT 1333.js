/// ==================== BOT 1333 CAR CENTER v8.0 ====================
// ✅ v8 ЗМІНИ:
//   • recalcClientBalance() — авто-перерахунок балансу з абонементів
//   • handleMarkVisit — списує кошти з абонементу + оновлює баланс клієнта
//   • confirmSubscriptionRequest — оновлює баланс клієнта після активації
//   • sell_ callback — оновлює баланс клієнта після продажу
//   • updateClientBalance() — єдина точка оновлення балансу
//   • syncAllBalances() — ручна функція для синхронізації всіх балансів

const scriptProperties = PropertiesService.getScriptProperties();
const BOT_TOKEN = scriptProperties.getProperty("BOT_TOKEN") || "";
const ADMIN_ID = Number(scriptProperties.getProperty("ADMIN_ID") || 0);
const SPREADSHEET_ID = scriptProperties.getProperty("SPREADSHEET_ID") || "";
const WEBAPP_URL = scriptProperties.getProperty("WEBAPP_URL") || "";
const PHONE = "+380 77 001 13 33";
const ADDRESS = "вул.Гвардійців-Широнінців 51/1";
const TZ = "Europe/Kyiv";
const CRM_SESSION_TTL_SECONDS = 21600;
const SCHEMA_VERSION = "audit-v7-2026-06-10";

const SHEET_CLIENTS      = "👥 Клієнти";
const SHEET_TYPES        = "💎 Типи абонементів";
const SHEET_SUBS         = "📋 Абонементи клієнтів";
const SHEET_PRICE        = "💰 Прайс";
const SHEET_SCHEDULE     = "📅 Розклад";
const SHEET_BOOKINGS     = "📝 Записи на послуги";
const SHEET_HISTORY      = "📜 Історія";
const SHEET_SUB_REQUESTS = "📬 Заявки на абонементи";
const SHEET_SUB_REQUESTS_ALTS = ["🔔 Заявки абонементів","Заявки абонементів"];
const SHEET_ERRORS       = "⚠️ Системные ошибки";
const SHEET_AUDIT_LOG    = "🧾 Журнал операций";
const SHEET_ANALYTICS    = "📊 Аналітика";

const DAY_NAMES       = ["Неділя","Понеділок","Вівторок","Середа","Четвер","П'ятниця","Субота"];
const DAY_NAMES_SHORT = ["Нд","Пн","Вт","Ср","Чт","Пт","Сб"];
const CLIENT_STATUS_ACTIVE = "Активный";
const CLIENT_STATUS_ARCHIVE = "Архив";
const CLIENT_STATUS_BLACKLIST = "Черный список";

const PUBLIC_ACTIONS = {
  getTimesAndBoxes: true,
  getAvailableTimesForDate: true,
  getBoxesForDateTime: true,
  saveBooking: true,
  cancelBooking: true,
  rescheduleBooking: true,
  saveSubOrder: true,
  autoRegister: true,
  verifyCrmPin: true
};

let __spreadsheet = null;
let __sheetCache = {};

// ─────────────────────── HELPERS ───────────────────────────

function requireConfig(name) {
  const value = scriptProperties.getProperty(name);
  if (value === null || value === undefined || String(value).trim() === "") {
    throw new Error("Missing Script Property: " + name);
  }
  return value;
}

function getSpreadsheet() {
  if (!__spreadsheet) __spreadsheet = SpreadsheetApp.openById(requireConfig("SPREADSHEET_ID"));
  return __spreadsheet;
}

function getSheet(name) {
  if (!name) return null;
  if (!__sheetCache[name]) {
    __sheetCache[name] = getSpreadsheet().getSheetByName(name);
  }
  return __sheetCache[name];
}

function safeGetRangeValues(sheet, startRow, startCol, numRows, numCols) {
  if (!sheet) return [];
  const last = sheet.getLastRow();
  if (last < startRow) return [];
  const rows = last - startRow + 1;
  if (rows <= 0) return [];
  const cols = numCols || Math.max(1, sheet.getLastColumn() - startCol + 1);
  return sheet.getRange(startRow, startCol, rows, cols).getValues();
}

function jsonResponse(result) {
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateSheet(name) {
  let sheet = getSheet(name);
  if (!sheet) {
    sheet = getSpreadsheet().insertSheet(name);
    __sheetCache[name] = sheet;
  }
  return sheet;
}

function ensureSheetHeaders(name, headers) {
  const sheet = getOrCreateSheet(name);
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const existing = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v || "").trim());
  const hasHeaders = existing.some(Boolean);
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (!hasHeaders) sheet.setFrozenRows(1);
  return sheet;
}

function ensureSchema(force) {
  if (!force && scriptProperties.getProperty("SCHEMA_VERSION") === SCHEMA_VERSION) return;
  ensureSheetHeaders(SHEET_CLIENTS, [
    "ClientID", "ПІБ", "Телефон", "Telegram ID", "Статус", "Дата реєстрації",
    "Баланс", "Количество посещений", "Общая сумма покупок", "Средний чек",
    "Последнее посещение", "Источник клиента", "Комментарий администратора"
  ]);
  ensureSheetHeaders(SHEET_SUBS, [
    "ClientID", "SubscriptionID", "TypeID", "Клієнт", "Тип", "Сума",
    "Дата покупки", "Дата окончания", "Витрачено", "Залишок", "Статус",
    "Кто продал", "Способ оплаты", "Скидка", "Промокод", "Дата активации"
  ]);
  ensureSheetHeaders(SHEET_BOOKINGS, [
    "BookingID", "Дата", "День", "Час", "Бокс", "ClientID", "Клієнт",
    "Телефон", "Послуга", "Статус", "Примітка", "Источник записи"
  ]);
  ensureSheetHeaders(SHEET_HISTORY, [
    "HistoryID", "Дата", "ClientID", "Клієнт", "SubscriptionID", "Операція",
    "Сума", "Коментар", "Кто выполнил операцию", "Изменение баланса", "До", "После"
  ]);
  ensureSheetHeaders(SHEET_SUB_REQUESTS, [
    "RequestID", "ClientID", "Клієнт", "Telegram ID", "TypeID", "Тип",
    "Сума", "Дата створення", "Статус", "Причина отказа", "Дата обработки", "Кто обработал"
  ]);
  ensureSheetHeaders(SHEET_ERRORS, ["Дата", "Пользователь", "Функция", "Ошибка", "StackTrace"]);
  ensureSheetHeaders(SHEET_AUDIT_LOG, ["Дата", "Пользователь", "Операция", "Сущность", "ID", "До", "После", "Комментарий"]);
  ensureSheetHeaders(SHEET_ANALYTICS, ["KPI", "Значение", "Обновлено"]);
  scriptProperties.setProperty("SCHEMA_VERSION", SCHEMA_VERSION);
}

function logError(functionName, error, user) {
  try {
    const sheet = ensureSheetHeaders(SHEET_ERRORS, ["Дата", "Пользователь", "Функция", "Ошибка", "StackTrace"]);
    sheet.appendRow([
      Utilities.formatDate(new Date(), TZ, "dd.MM.yyyy HH:mm:ss"),
      String(user || ""),
      String(functionName || ""),
      String(error && error.message ? error.message : error),
      String(error && error.stack ? error.stack : "")
    ]);
  } catch (logErr) {
    Logger.log("logError failed: " + logErr);
  }
}

function logAction(user, operation, entity, entityId, beforeValue, afterValue, comment) {
  try {
    const sheet = ensureSheetHeaders(SHEET_AUDIT_LOG, ["Дата", "Пользователь", "Операция", "Сущность", "ID", "До", "После", "Комментарий"]);
    sheet.appendRow([
      Utilities.formatDate(new Date(), TZ, "dd.MM.yyyy HH:mm:ss"),
      String(user || ""),
      String(operation || ""),
      String(entity || ""),
      String(entityId || ""),
      beforeValue === undefined ? "" : String(beforeValue),
      afterValue === undefined ? "" : String(afterValue),
      String(comment || "")
    ]);
  } catch (e) {
    logError("logAction", e, user);
  }
}

function withScriptLock(functionName, fn, user) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    return fn();
  } catch (e) {
    logError(functionName, e, user);
    throw e;
  } finally {
    try { lock.releaseLock(); } catch (releaseErr) {}
  }
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
  let digits = String(phone).replace(/\D/g, "");
  // Нормалізуємо до 12 цифр (380XXXXXXXXX)
  if (digits.startsWith("380") && digits.length === 12) {
    // вже OK
  } else if (digits.startsWith("0") && digits.length === 10) {
    digits = "38" + digits;
  } else if (digits.length === 9) {
    digits = "380" + digits;
  } else if (digits.startsWith("38") && digits.length === 11) {
    digits = "3" + digits;
  }
  // Формат: +38(0XX)XXX XX XX
  if (digits.length === 12 && digits.startsWith("38")) {
    return `+38(${digits.slice(2,5)})${digits.slice(5,8)} ${digits.slice(8,10)} ${digits.slice(10,12)}`;
  }
  return phone.startsWith("+") ? phone : "+" + digits;
}

function sha256Hex(value) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(value), Utilities.Charset.UTF_8)
    .map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2, "0"))
    .join("");
}

function secureCompare(a, b) {
  a = String(a || "");
  b = String(b || "");
  if (a.length !== b.length) return false;
  let diff = 0;
  for (let i = 0; i < a.length; i++) diff |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return diff === 0;
}

function isCrmPinValid(pin) {
  const hash = scriptProperties.getProperty("CRM_PIN_HASH");
  if (hash) return secureCompare(sha256Hex(pin), String(hash).toLowerCase());
  const plain = scriptProperties.getProperty("CRM_PIN");
  return !!plain && secureCompare(String(pin), String(plain));
}

function createAdminSession() {
  const token = Utilities.getUuid() + "-" + Utilities.getUuid();
  CacheService.getScriptCache().put("crm_session_" + token, "1", CRM_SESSION_TTL_SECONDS);
  return token;
}

function isValidAdminToken(token) {
  if (!token) return false;
  return CacheService.getScriptCache().get("crm_session_" + token) === "1";
}

function verifyCrmPin(pin) {
  if (!isCrmPinValid(pin)) {
    logAction("web", "Неверный PIN", "CRM", "", "", "", "Попытка входа");
    return { success: false, error: "Невірний PIN" };
  }
  const token = createAdminSession();
  logAction("web", "Вход в CRM", "CRM", "", "", "ok", "Серверная авторизация");
  return { success: true, token, expiresIn: CRM_SESSION_TTL_SECONDS };
}

function requiresAdminAuth(action) {
  return action && !PUBLIC_ACTIONS[action];
}

function invalidateBookingCache(dateStr) {
  try {
    const cache = CacheService.getScriptCache();
    if (dateStr) cache.remove("timesBoxes_" + dateStr);
    if (dateStr) cache.remove("avail_" + dateStr);
  } catch (e) {}
}

function asDate(value) {
  if (!value) return null;
  if (value instanceof Date) return value;
  if (typeof value === "number") {
    const dt = new Date((Math.floor(value) - 25569) * 86400 * 1000);
    return isNaN(dt.getTime()) ? null : dt;
  }
  const parts = String(value).trim().match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
  if (!parts) return null;
  const dt = new Date(Number(parts[3]), Number(parts[2]) - 1, Number(parts[1]));
  return isNaN(dt.getTime()) ? null : dt;
}

function isPastDate(value) {
  const dt = asDate(value);
  if (!dt) return false;
  const today = new Date();
  const todayOnly = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  const dateOnly = new Date(dt.getFullYear(), dt.getMonth(), dt.getDate());
  return dateOnly < todayOnly;
}

function isActiveStatus(status) {
  const s = String(status || "");
  return s.includes("Активний") || s.includes("Активный") || s.includes("✅");
}

function isClientBlockedStatus(status) {
  const s = String(status || "");
  return s.includes(CLIENT_STATUS_BLACKLIST) || s.includes(CLIENT_STATUS_ARCHIVE);
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
  const subsData = safeGetRangeValues(subsSheet, 2, 1, null, 16);
  let totalBalance = 0;
  
  subsData.forEach(r => {
    if (!r[0]) return;
    if (String(r[0]).trim() !== String(clientId).trim()) return;
    const status = String(r[10]).trim();
    // Враховуємо тільки активні абонементи
    if (isActiveStatus(status)) {
      const bal = parseFloat(r[9]) || 0;
      totalBalance += bal;
    }
  });
  
  // Записуємо в колонку G (індекс 7) листа Клієнти
  const clientsData = safeGetRangeValues(clientSheet, 2, 1, null, 1);
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
  
  const clientsData = safeGetRangeValues(clientSheet, 2, 1, null, 1);
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
  
  const data = safeGetRangeValues(sheet, 2, 1, null, 16);
  
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
    const newStatus = newBalance <= 0 ? "⛔ Вичерпано" : data[i][10];
    sheet.getRange(i + 2, 9, 1, 3).setValues([[newSpent, newBalance, newStatus]]);
    
    Logger.log(`✅ Списано ${cost} грн з ${subId}. Залишок: ${newBalance}`);
    
    // Оновлюємо загальний баланс клієнта
    updateClientBalance(clientId);
    return true;
  }
  
  Logger.log(`❌ Абонемент не знайдено: ${subId}`);
  return false;
}

function findSubForCost(clientId, cost) {
  if (!clientId || !cost || cost <= 0) return null;
  const subs = getClientSubs(clientId)
    .filter(s => isActiveStatus(s.status) && parseFloat(s.balance) >= cost);
  if (!subs.length) return null;
  subs.sort((a, b) => parseFloat(b.balance) - parseFloat(a.balance));
  return subs[0].subId;
}

function markVisit(clientId, subId, service, cost, note, actor) {
  return withScriptLock("markVisit", function() {
    assertForeignKeys({ clientId, subId });
    const client = findClientById(clientId);
    const beforeBalance = getClientLiveBalance(clientId);
    const deducted = deductFromSub(clientId, subId, cost);
    if (!deducted) return { success: false, error: "Недостатньо коштів" };
    const afterBalance = getClientLiveBalance(clientId);
    addToHistory("", clientId, client?.name || "", fmtDate(new Date()), service, cost, note, actor, afterBalance - beforeBalance, beforeBalance, afterBalance, subId);
    updateClientMetrics(clientId);
    logAction(actor || "admin", "Списание", "Subscription", subId, beforeBalance, afterBalance, service);
    return { success: true, beforeBalance, afterBalance };
  }, actor || "admin");
}

// ──── AUTO-REGISTER CLIENT ────────────────────────────────────
function autoRegisterClient(name, phone, telegramId, source) {
  if (!name || !phone) return null;
  return withScriptLock("autoRegisterClient", function() {
    const normPhone = normalizePhone(phone);
    const tg = String(telegramId || "").trim();
    const sheet = getSheet(SHEET_CLIENTS);
    if (!sheet) return null;

    const data = safeGetRangeValues(sheet, 2, 1, null, 13);
    for (let i = 0; i < data.length; i++) {
      if (!data[i][0]) continue;
      const samePhone = normPhone && normalizePhone(String(data[i][2])) === normPhone;
      const sameTelegram = tg && String(data[i][3]).trim() === tg;
      if (samePhone || sameTelegram) {
        const row = i + 2;
        const updates = data[i].slice(0, 13);
        if (!updates[2] && normPhone) updates[2] = normPhone;
        if (!updates[3] && tg) updates[3] = tg;
        if (!updates[4]) updates[4] = CLIENT_STATUS_ACTIVE;
        if (!updates[11] && source) updates[11] = source;
        sheet.getRange(row, 1, 1, updates.length).setValues([updates]);
        Logger.log("✅ Клієнт вже є: " + data[i][0]);
        return String(data[i][0]);
      }
    }

    const row = getNextRow(SHEET_CLIENTS);
    const clientId = genClientId();
    sheet.getRange(row, 1, 1, 13).setValues([[
      clientId, name, normPhone, tg, CLIENT_STATUS_ACTIVE, fmtDate(new Date()), 0,
      0, 0, 0, "", source || "Сайт", ""
    ]]);

    Logger.log("🆕 Авто-реєстрація: " + clientId + " " + name);
    logAction(source || "web", "Создание клиента", "Client", clientId, "", name, normPhone);

    sendMessage(ADMIN_ID,
      `👤 <b>НОВИЙ КЛІЄНТ</b>\n\n` +
      `🆔 ${clientId}\n👤 ${name}\n📱 ${normPhone}\n📌 ${source || "Сайт"}`
    );

    return clientId;
  }, source || "web");
}

// ──── LINK TELEGRAM ID ────────────────────────────────────────
function linkTelegramId(phone, telegramId) {
  const normPhone = normalizePhone(phone);
  const sheet = getSheet(SHEET_CLIENTS);
  if (!sheet) return false;
  
  const data = safeGetRangeValues(sheet, 2, 1, null, 13);
  
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
  const data = safeGetRangeValues(getSheet(SHEET_CLIENTS), 2, 1, null, 1);
  let max = 0;
  data.forEach(r => { if (r[0]) { const m = String(r[0]).match(/C-(\d+)/); if (m) max = Math.max(max, +m[1]); }});
  return "C-" + String(max + 1).padStart(3, "0");
}
function genBookingId() {
  const data = safeGetRangeValues(getSheet(SHEET_BOOKINGS), 2, 1, null, 1);
  let max = 0;
  data.forEach(r => { if (r[0]) { const m = String(r[0]).match(/B-(\d+)/); if (m) max = Math.max(max, +m[1]); }});
  return "B-" + String(max + 1).padStart(4, "0");
}
function genSubId() {
  const data = safeGetRangeValues(getSheet(SHEET_SUBS), 2, 1, null, 2);
  let max = 0;
  data.forEach(r => { if (r[1]) { const m = String(r[1]).match(/A-(\d+)/); if (m) max = Math.max(max, +m[1]); }});
  return "A-" + String(max + 1).padStart(3, "0");
}
function genHistoryId() {
  const data = safeGetRangeValues(getSheet(SHEET_HISTORY), 2, 1, null, 1);
  let max = 0;
  data.forEach(r => { if (r[0]) { const m = String(r[0]).match(/H-(\d+)/); if (m) max = Math.max(max, +m[1]); }});
  return "H-" + String(max + 1).padStart(4, "0");
}
function genSubRequestId() {
  const sheet = getSubRequestSheet();
  if (!sheet) return "SR-0001";
  const data = safeGetRangeValues(sheet, 2, 1, null, 1);
  let max = 0;
  data.forEach(r => { if (r[0]) { const m = String(r[0]).match(/SR-(\d+)/); if (m) max = Math.max(max, +m[1]); }});
  return "SR-" + String(max + 1).padStart(4, "0");
}

function getNextRow(sheetName) {
  const sheet = getSheet(sheetName);
  if (!sheet) return 2;
  const data = safeGetRangeValues(sheet, 2, 1, null, 1);
  for (let i = 0; i < data.length; i++) { if (!data[i][0]) return i + 2; }
  return sheet.getLastRow() + 1;
}

// ─────────────────────── TELEGRAM API ──────────────────────────

function sendMessage(chatId, text, keyboard = null) {
  if (!BOT_TOKEN) {
    logError("sendMessage", new Error("Missing BOT_TOKEN Script Property"), chatId);
    return null;
  }
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
    if (!BOT_TOKEN) throw new Error("Missing BOT_TOKEN Script Property");
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
  const data = safeGetRangeValues(getSheet(SHEET_CLIENTS), 2, 1, null, 13);
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] && String(data[i][3]).trim() === String(telegramId).trim()) {
      // ✅ Беремо актуальний баланс з абонементів, а не з кешу в таблиці
      const liveBalance = getClientLiveBalance(String(data[i][0]));
      return {
        clientId: data[i][0], name: data[i][1], phone: data[i][2],
        telegramId: String(data[i][3] || '').trim(),
        status: String(data[i][4] || CLIENT_STATUS_ACTIVE),
        registered: parseDate(data[i][5]), balance: liveBalance,
        visits: Number(data[i][7]) || 0,
        totalPurchases: Number(data[i][8]) || 0,
        averageCheck: Number(data[i][9]) || 0,
        lastVisit: parseDate(data[i][10]),
        source: String(data[i][11] || ""),
        adminComment: String(data[i][12] || "")
      };
    }
  }
  return null;
}

function findClientById(clientId) {
  const data = safeGetRangeValues(getSheet(SHEET_CLIENTS), 2, 1, null, 13);
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(clientId).trim()) {
      const liveBalance = getClientLiveBalance(String(data[i][0]));
      return {
        clientId: data[i][0], name: data[i][1], phone: data[i][2],
        telegramId: String(data[i][3] || '').trim(),
        status: String(data[i][4] || CLIENT_STATUS_ACTIVE),
        registered: parseDate(data[i][5]), balance: liveBalance,
        visits: Number(data[i][7]) || 0,
        totalPurchases: Number(data[i][8]) || 0,
        averageCheck: Number(data[i][9]) || 0,
        lastVisit: parseDate(data[i][10]),
        source: String(data[i][11] || ""),
        adminComment: String(data[i][12] || "")
      };
    }
  }
  return null;
}

function findClientByPhone(phone) {
  const normPhone = normalizePhone(phone);
  const data = safeGetRangeValues(getSheet(SHEET_CLIENTS), 2, 1, null, 13);
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] && normalizePhone(String(data[i][2])) === normPhone) {
      const liveBalance = getClientLiveBalance(String(data[i][0]));
      return {
        row: i + 2, clientId: data[i][0], name: data[i][1], phone: data[i][2],
        telegramId: String(data[i][3]).trim(),
        status: String(data[i][4] || CLIENT_STATUS_ACTIVE),
        balance: liveBalance
      };
    }
  }
  return null;
}

// ✅ Живий баланс клієнта — рахується з абонементів в реальному часі
function getClientLiveBalance(clientId) {
  if (!clientId) return 0;
  const subsData = safeGetRangeValues(getSheet(SHEET_SUBS), 2, 1, null, 16);
  let total = 0;
  subsData.forEach(r => {
    if (!r[0]) return;
    if (String(r[0]).trim() !== String(clientId).trim()) return;
    const status = String(r[10]).trim();
    if (isActiveStatus(status)) {
      total += parseFloat(r[9]) || 0;
    }
  });
  return total;
}

// ─────────────────────── DATA GETTERS ──────────────────────────

function getTypes() {
  const data = safeGetRangeValues(getSheet(SHEET_TYPES), 2, 1, null, 6);
  return data.filter(r => r[0]).map(r => ({
    id: r[0], name: r[1], months: r[2], price: r[3], bonus: r[4], amount: r[5]
  }));
}

function getClientSubs(clientId) {
  const data = safeGetRangeValues(getSheet(SHEET_SUBS), 2, 1, null, 16);
  return data
    .map((r, i) => ({ ...r, _row: i + 2 }))
    .filter(r => String(r[0]).trim() === String(clientId).trim() && r[1])
    .map(r => ({
      row: r._row, subId: r[1], typeName: r[4], amount: r[5],
      startDate: parseDate(r[6]), endDate: parseDate(r[7]),
      spent: parseFloat(r[8]) || 0, balance: parseFloat(r[9]) || 0,
      status: r[10], soldBy: r[11] || "", paymentMethod: r[12] || "",
      discount: r[13] || "", promoCode: r[14] || "", activationDate: parseDate(r[15])
    }));
}

function getClientHistory(clientId) {
  const data = safeGetRangeValues(getSheet(SHEET_HISTORY), 2, 1, null, 12);
  return data.filter(r => String(r[2]).trim() === String(clientId).trim())
    .map(r => ({
      date: parseDate(r[1]), service: r[5], cost: r[6], note: r[7],
      actor: r[8] || "", balanceChange: r[9] || "", before: r[10] || "", after: r[11] || ""
    }));
}

function getClientBookings(clientId) {
  const sheet = getSheet(SHEET_BOOKINGS);
  if (!sheet) return [];
  const data = safeGetRangeValues(sheet, 2, 1, null, 12);
  return data.filter(r => r[0] && String(r[5]).trim() === String(clientId).trim()).map(row => ({
    id: String(row[0]),
    date: row[1] instanceof Date ? fmtDate(row[1]) : String(row[1]).trim(),
    time: row[3] instanceof Date ? Utilities.formatDate(row[3], TZ, "HH:mm") : String(row[3]).trim(),
    box: String(row[4]), service: String(row[8]), status: String(row[9]), note: String(row[10] || ""), source: String(row[11] || "")
  }));
}

function getPriceByType(carType) {
  const data = safeGetRangeValues(getSheet(SHEET_PRICE), 2, 1, null, 4);
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
  const data = safeGetRangeValues(sheet, 2, 1, null, 8);
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

function getBoxReleaseText(dateStr, boxName) {
  const sheet = getSheet(SHEET_BOOKINGS);
  const data = safeGetRangeValues(sheet, 2, 1, null, 12);
  const now = new Date();
  const nowMinutes = Number(Utilities.formatDate(now, TZ, "H")) * 60 + Number(Utilities.formatDate(now, TZ, "m"));
  let best = null;
  data.forEach(r => {
    if (!r[0]) return;
    const bDate = r[1] instanceof Date ? fmtDate(r[1]) : String(r[1]).trim();
    const bBox = String(r[4]).trim();
    const status = String(r[9]).trim();
    if (bDate !== dateStr || bBox !== boxName) return;
    if (!(status.includes("роботі") || status.includes("Підтверджено") || status.includes("Очікує"))) return;
    const time = r[3] instanceof Date ? Utilities.formatDate(r[3], TZ, "HH:mm") : normalizeTime(r[3]);
    const h = parseTimeString(time);
    if (h === null) return;
    const releaseMinutes = (h + 1) * 60;
    if (releaseMinutes >= nowMinutes && (best === null || releaseMinutes < best)) best = releaseMinutes;
  });
  if (best === null) return "";
  const diff = best - nowMinutes;
  if (diff <= 0) return "звільняється зараз";
  return "через " + Math.floor(diff / 60) + " год " + String(diff % 60).padStart(2, "0") + " хв";
}

function getClientRow(clientId) {
  const data = safeGetRangeValues(getSheet(SHEET_CLIENTS), 2, 1, null, 13);
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(clientId).trim()) {
      return { row: i + 2, values: data[i] };
    }
  }
  return null;
}

function clientExists(clientId) {
  return !!getClientRow(clientId);
}

function typeExists(typeId) {
  return getTypes().some(t => String(t.id).trim() === String(typeId).trim());
}

function subscriptionExists(subId, clientId) {
  const data = safeGetRangeValues(getSheet(SHEET_SUBS), 2, 1, null, 16);
  return data.some(r =>
    String(r[1]).trim() === String(subId).trim() &&
    (!clientId || String(r[0]).trim() === String(clientId).trim())
  );
}

function assertClientAllowed(clientId) {
  const found = getClientRow(clientId);
  if (!found) throw new Error("ClientID not found: " + clientId);
  const status = found.values[4] || CLIENT_STATUS_ACTIVE;
  if (isClientBlockedStatus(status)) throw new Error("Client is blocked/archived: " + clientId);
  return found;
}

function assertForeignKeys(keys) {
  if (keys.clientId) assertClientAllowed(keys.clientId);
  if (keys.typeId && !typeExists(keys.typeId)) throw new Error("TypeID not found: " + keys.typeId);
  if (keys.subId && !subscriptionExists(keys.subId, keys.clientId)) throw new Error("SubscriptionID not found: " + keys.subId);
}

function updateClientMetrics(clientId) {
  const client = getClientRow(clientId);
  if (!client) return false;
  const histData = safeGetRangeValues(getSheet(SHEET_HISTORY), 2, 1, null, 12);
  const subData = safeGetRangeValues(getSheet(SHEET_SUBS), 2, 1, null, 16);
  let visits = 0;
  let visitTotal = 0;
  let lastVisit = null;
  histData.forEach(r => {
    if (!r[0] || String(r[2]).trim() !== String(clientId).trim()) return;
    visits++;
    visitTotal += parseFloat(r[6]) || 0;
    const d = asDate(r[1]);
    if (d && (!lastVisit || d > lastVisit)) lastVisit = d;
  });
  let totalPurchases = 0;
  subData.forEach(r => {
    if (!r[0] || String(r[0]).trim() !== String(clientId).trim()) return;
    totalPurchases += parseFloat(r[5]) || 0;
  });
  const averageCheck = visits ? visitTotal / visits : 0;
  const rowValues = client.values.slice(0, 13);
  rowValues[4] = rowValues[4] || CLIENT_STATUS_ACTIVE;
  rowValues[6] = getClientLiveBalance(clientId);
  rowValues[7] = visits;
  rowValues[8] = totalPurchases;
  rowValues[9] = fmt(averageCheck);
  rowValues[10] = lastVisit ? fmtDate(lastVisit) : "";
  getSheet(SHEET_CLIENTS).getRange(client.row, 1, 1, rowValues.length).setValues([rowValues]);
  return true;
}

function syncClientMetrics() {
  const data = safeGetRangeValues(getSheet(SHEET_CLIENTS), 2, 1, null, 13);
  let count = 0;
  data.forEach(r => {
    if (r[0] && updateClientMetrics(String(r[0]).trim())) count++;
  });
  return count;
}

function expireSubscriptions() {
  return withScriptLock("expireSubscriptions", function() {
    const sheet = getSheet(SHEET_SUBS);
    const data = safeGetRangeValues(sheet, 2, 1, null, 16);
    let count = 0;
    data.forEach((r, i) => {
      if (!r[0] || !r[1]) return;
      const status = String(r[10] || "");
      if (!isActiveStatus(status)) return;
      if (!isPastDate(r[7])) return;
      sheet.getRange(i + 2, 11).setValue("⛔ Просрочен");
      updateClientBalance(String(r[0]));
      updateClientMetrics(String(r[0]));
      logAction("system", "Автозакрытие абонемента", "Subscription", r[1], status, "⛔ Просрочен", parseDate(r[7]));
      count++;
    });
    return count;
  }, "system");
}

function createDailyBackup() {
  try {
    const backupName = "Backup_" + Utilities.formatDate(new Date(), TZ, "yyyy_MM_dd");
    const file = DriveApp.getFileById(requireConfig("SPREADSHEET_ID"));
    const parents = file.getParents();
    const folder = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
    const existing = folder.getFilesByName(backupName);
    if (existing.hasNext()) return { created: false, name: backupName };
    file.makeCopy(backupName, folder);
    logAction("system", "Backup", "Spreadsheet", backupName, "", "created", "");
    return { created: true, name: backupName };
  } catch (e) {
    // Помилка дозволів Drive або інша проблема — залогуємо і продовжуємо
    Logger.log("⚠️ Резервна копія не створена: " + e.message);
    return { created: false, error: e.message };
  }
}

function mergeDuplicateClients() {
  return withScriptLock("mergeDuplicateClients", function() {
    const clientSheet = getSheet(SHEET_CLIENTS);
    const data = safeGetRangeValues(clientSheet, 2, 1, null, 13);
    const seen = {};
    let merged = 0;
    data.forEach((r, i) => {
      if (!r[0]) return;
      const keys = [];
      const phone = normalizePhone(String(r[2] || ""));
      const tg = String(r[3] || "").trim();
      if (phone) keys.push("p:" + phone);
      if (tg) keys.push("t:" + tg);
      let primary = null;
      keys.forEach(k => { if (!primary && seen[k]) primary = seen[k]; });
      if (!primary) {
        keys.forEach(k => seen[k] = String(r[0]));
        return;
      }
      const duplicateId = String(r[0]);
      if (duplicateId === primary) return;
      replaceClientId(duplicateId, primary);
      const rowValues = r.slice(0, 13);
      rowValues[4] = CLIENT_STATUS_ARCHIVE;
      rowValues[12] = "Merged into " + primary + " " + Utilities.formatDate(new Date(), TZ, "dd.MM.yyyy HH:mm");
      clientSheet.getRange(i + 2, 1, 1, rowValues.length).setValues([rowValues]);
      updateClientMetrics(primary);
      logAction("system", "Объединение дубля", "Client", duplicateId, duplicateId, primary, phone || tg);
      merged++;
    });
    return merged;
  }, "system");
}

function replaceClientId(oldId, newId) {
  const updates = [
    { sheet: getSheet(SHEET_SUBS), col: 1 },
    { sheet: getSheet(SHEET_BOOKINGS), col: 6 },
    { sheet: getSheet(SHEET_HISTORY), col: 3 },
    { sheet: getSubRequestSheet(), col: 2 }
  ];
  updates.forEach(cfg => {
    if (!cfg.sheet) return;
    const data = safeGetRangeValues(cfg.sheet, 2, cfg.col, null, 1);
    data.forEach((r, i) => {
      if (String(r[0]).trim() === String(oldId).trim()) {
        cfg.sheet.getRange(i + 2, cfg.col).setValue(newId);
      }
    });
  });
}

function calculateAnalytics() {
  const clients = safeGetRangeValues(getSheet(SHEET_CLIENTS), 2, 1, null, 13).filter(r => r[0]);
  const subs = safeGetRangeValues(getSheet(SHEET_SUBS), 2, 1, null, 16).filter(r => r[0]);
  const history = safeGetRangeValues(getSheet(SHEET_HISTORY), 2, 1, null, 12).filter(r => r[0]);
  const bookings = safeGetRangeValues(getSheet(SHEET_BOOKINGS), 2, 1, null, 12).filter(r => r[0]);
  const todayStr = Utilities.formatDate(new Date(), TZ, "dd.MM.yyyy");
  const now = new Date();
  const weekStart = new Date(now);
  weekStart.setDate(now.getDate() - 7);
  const monthKey = Utilities.formatDate(now, TZ, "MM.yyyy");

  const saleRows = subs.map(r => ({ amount: parseFloat(r[5]) || 0, date: asDate(r[6]), service: String(r[4] || ""), clientId: String(r[0]) }));
  const salesToday = saleRows.filter(x => x.date && fmtDate(x.date) === todayStr).reduce((s, x) => s + x.amount, 0);
  const salesWeek = saleRows.filter(x => x.date && x.date >= weekStart).reduce((s, x) => s + x.amount, 0);
  const salesMonth = saleRows.filter(x => x.date && Utilities.formatDate(x.date, TZ, "MM.yyyy") === monthKey).reduce((s, x) => s + x.amount, 0);
  const totalRevenue = saleRows.reduce((s, x) => s + x.amount, 0) + history.reduce((s, r) => s + (parseFloat(r[6]) || 0), 0);
  const orderCount = saleRows.length + history.length;
  const serviceRevenue = history.reduce((s, r) => s + (parseFloat(r[6]) || 0), 0);
  const subRevenue = saleRows.reduce((s, x) => s + x.amount, 0);
  const serviceCounts = {};
  history.forEach(r => { const key = String(r[5] || "—"); serviceCounts[key] = (serviceCounts[key] || 0) + 1; });
  bookings.forEach(r => { const key = String(r[8] || "—").split("|")[0].trim(); serviceCounts[key] = (serviceCounts[key] || 0) + 1; });
  const popularService = Object.keys(serviceCounts).sort((a, b) => serviceCounts[b] - serviceCounts[a])[0] || "—";
  const carTypes = {};
  bookings.forEach(r => {
    const match = String(r[8] || "").match(/\|\s*([^|]+)$/);
    const key = match ? match[1].trim() : "—";
    carTypes[key] = (carTypes[key] || 0) + 1;
  });
  const popularCar = Object.keys(carTypes).sort((a, b) => carTypes[b] - carTypes[a])[0] || "—";
  const clientRevenue = {};
  saleRows.forEach(x => clientRevenue[x.clientId] = (clientRevenue[x.clientId] || 0) + x.amount);
  history.forEach(r => { const cid = String(r[2] || ""); clientRevenue[cid] = (clientRevenue[cid] || 0) + (parseFloat(r[6]) || 0); });
  const topClientId = Object.keys(clientRevenue).sort((a, b) => clientRevenue[b] - clientRevenue[a])[0] || "";
  const topClient = clients.find(r => String(r[0]) === topClientId);
  const todayBookings = bookings.filter(r => (r[1] instanceof Date ? fmtDate(r[1]) : String(r[1]).trim()) === todayStr);
  const busyBoxes = todayBookings.filter(r => {
    const s = String(r[9] || "");
    return s.includes("роботі") || s.includes("Підтверджено") || s.includes("Очікує");
  }).length;
  const activeClients = new Set(bookings.map(r => String(r[5] || "")).concat(history.map(r => String(r[2] || "")))).size;
  const repeatClients = clients.filter(c => (Number(c[7]) || 0) > 1).length;

  return [
    ["Количество клиентов", clients.length],
    ["Активные абонементы", subs.filter(r => isActiveStatus(r[10])).length],
    ["Продажи сегодня", salesToday],
    ["Продажи недели", salesWeek],
    ["Продажи месяца", salesMonth],
    ["Средний чек", orderCount ? fmt(totalRevenue / orderCount) : 0],
    ["Доход по услугам", serviceRevenue],
    ["Доход по абонементам", subRevenue],
    ["Самая популярная услуга", popularService],
    ["Самый популярный тип авто", popularCar],
    ["Самый прибыльный клиент", topClient ? `${topClient[1]} (${topClientId})` : "—"],
    ["Загрузка боксов", `${busyBoxes}/3`],
    ["Процент повторных клиентов", clients.length ? fmt((repeatClients / clients.length) * 100) + "%" : "0%"],
    ["Клиентов с активностью", activeClients]
  ];
}

function updateAnalyticsSheet() {
  const sheet = ensureSheetHeaders(SHEET_ANALYTICS, ["KPI", "Значение", "Обновлено"]);
  const updated = Utilities.formatDate(new Date(), TZ, "dd.MM.yyyy HH:mm:ss");
  const rows = calculateAnalytics().map(r => [r[0], r[1], updated]);
  if (sheet.getLastRow() > 1) sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).clearContent();
  if (rows.length) sheet.getRange(2, 1, rows.length, 3).setValues(rows);
  return rows;
}

function sendVisitReminders() {
  const data = safeGetRangeValues(getSheet(SHEET_BOOKINGS), 2, 1, null, 12);
  let sent = 0;
  data.forEach(r => {
    if (!r[0] || !r[5]) return;
    const status = String(r[9] || "");
    if (!(status.includes("Підтверджено") || status.includes("Очікує"))) return;
    const dt = asDate(r[1]);
    const hour = parseTimeString(r[3]);
    if (!dt || hour === null) return;
    dt.setHours(hour, 0, 0, 0);
    const diffHours = (dt.getTime() - new Date().getTime()) / 3600000;
    [[24, "24h"], [1, "1h"]].forEach(([target, label]) => {
      if (diffHours < target || diffHours > target + 1) return;
      const key = `reminder_visit_${r[0]}_${label}`;
      if (scriptProperties.getProperty(key)) return;
      const client = findClientById(String(r[5]));
      if (client?.telegramId) {
        sendMessage(client.telegramId, `⏰ <b>Нагадування про запис</b>\n\n📅 ${fmtDate(dt)} ${normalizeTime(r[3])}\n🏠 ${r[4]}\n🔧 ${r[8]}`);
        scriptProperties.setProperty(key, "1");
        sent++;
      }
    });
  });
  return sent;
}

function sendSubscriptionExpiryReminders() {
  const data = safeGetRangeValues(getSheet(SHEET_SUBS), 2, 1, null, 16);
  let sent = 0;
  data.forEach(r => {
    if (!r[0] || !r[1] || !isActiveStatus(r[10])) return;
    const end = asDate(r[7]);
    if (!end) return;
    const today = new Date();
    const days = Math.ceil((new Date(end.getFullYear(), end.getMonth(), end.getDate()) - new Date(today.getFullYear(), today.getMonth(), today.getDate())) / 86400000);
    if ([7, 3, 1].indexOf(days) === -1) return;
    const key = `reminder_sub_${r[1]}_${days}`;
    if (scriptProperties.getProperty(key)) return;
    const client = findClientById(String(r[0]));
    if (client?.telegramId) {
      sendMessage(client.telegramId, `💎 <b>Абонемент скоро закінчиться</b>\n\n${r[4]}\n📅 До: ${parseDate(r[7])}\n⏳ Залишилось: ${days} дн.`);
      scriptProperties.setProperty(key, "1");
      sent++;
    }
  });
  return sent;
}

function broadcastTelegram(message, segment, actor) {
  if (!message) return { success: false, sent: 0, error: "message required" };
  const clients = safeGetRangeValues(getSheet(SHEET_CLIENTS), 2, 1, null, 13).filter(r => r[0] && r[3]);
  let sent = 0;
  clients.forEach(r => {
    const status = String(r[4] || CLIENT_STATUS_ACTIVE);
    if (isClientBlockedStatus(status)) return;
    if (segment === "vip" && !status.includes("VIP")) return;
    if (sendMessage(String(r[3]).trim(), message)) sent++;
  });
  logAction(actor || "CRM", "Массовая рассылка", "Telegram", segment || "all", "", sent, "");
  return { success: true, sent };
}

function runDailyMaintenance(force) {
  const key = Utilities.formatDate(new Date(), TZ, "yyyy_MM_dd");
  if (!force && scriptProperties.getProperty("DAILY_MAINTENANCE_DATE") === key) return { skipped: true };
  const expired = expireSubscriptions();
  const merged = mergeDuplicateClients();
  const balances = syncAllBalances();
  const metrics = syncClientMetrics();
  const analytics = updateAnalyticsSheet().length;
  let backup = { created: false };
  try { backup = createDailyBackup(); } catch (e) { logError("createDailyBackup", e, "system"); }
  scriptProperties.setProperty("DAILY_MAINTENANCE_DATE", key);
  return { skipped: false, expired, merged, balances, metrics, analytics, backup };
}

function runHourlyReminders() {
  return {
    visits: sendVisitReminders(),
    subscriptions: sendSubscriptionExpiryReminders()
  };
}

function setupProduction() {
  ensureSchema(true);
  installProductionTriggers();
  return runDailyMaintenance(true);
}

function installProductionTriggers() {
  const managed = ["runDailyMaintenance", "runHourlyReminders"];
  ScriptApp.getProjectTriggers().forEach(t => {
    if (managed.indexOf(t.getHandlerFunction()) !== -1) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("runDailyMaintenance").timeBased().everyDays(1).atHour(3).create();
  ScriptApp.newTrigger("runHourlyReminders").timeBased().everyHours(1).create();
  return true;
}

// ─────────────────────── BOOKING ───────────────────────────────

function isBoxBooked(dateStr, timeStr, boxName) {
  const sheet = getSheet(SHEET_BOOKINGS);
  if (!sheet) return false;
  const data = safeGetRangeValues(sheet, 2, 1, null, 12);
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
    if (bDate === dateStr && bBox === boxName &&
        (bStatus.includes("Підтверджено") || bStatus.includes("Очікує") || bStatus.includes("роботі"))) {
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

function saveBooking(dateStr, timeStr, boxName, clientId, clientName, phone, service, source, note, actor) {
  return withScriptLock("saveBooking", function() {
    const sheet = getSheet(SHEET_BOOKINGS);
    if (!sheet) return false;
    assertForeignKeys({ clientId });
    if (isBoxBooked(dateStr, timeStr, boxName)) {
      Logger.log("❌ Бокс вже зайнятий: " + boxName);
      return false;
    }
    const row = getNextRow(SHEET_BOOKINGS);
    const [day, month, year] = dateStr.split(".");
    const d = new Date(parseInt(year), parseInt(month) - 1, parseInt(day));
    const dayName = DAY_NAMES[d.getDay()];
    const normTime = normalizeTime(timeStr);
    const bookingId = genBookingId();
    sheet.getRange(row, 1, 1, 12).setValues([[
      bookingId, dateStr, dayName, normTime, boxName, clientId, clientName,
      normalizePhone(phone), service, "✅ Підтверджено", note || "", source || "Администратор"
    ]]);
    invalidateBookingCache(dateStr);
    logAction(actor || source || "web", "Запись", "Booking", bookingId, "", "✅ Підтверджено", `${dateStr} ${normTime} ${boxName}`);
    return bookingId;
  }, actor || source || "web");
}

function findBookingById(bookingId) {
  const sheet = getSheet(SHEET_BOOKINGS);
  const data = safeGetRangeValues(sheet, 2, 1, null, 12);
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(bookingId).trim()) {
      return { row: i + 2, values: data[i] };
    }
  }
  return null;
}

function cancelBooking(bookingId, phone, actor) {
  return withScriptLock("cancelBooking", function() {
    const booking = findBookingById(bookingId);
    if (!booking) return { success: false, error: "booking not found" };
    const bookingPhone = normalizePhone(String(booking.values[7] || ""));
    if (phone && bookingPhone !== normalizePhone(phone)) return { success: false, error: "phone mismatch" };
    const oldStatus = String(booking.values[9] || "");
    getSheet(SHEET_BOOKINGS).getRange(booking.row, 10).setValue("❌ Скасовано клієнтом");
    const dateStr = booking.values[1] instanceof Date ? fmtDate(booking.values[1]) : String(booking.values[1]).trim();
    invalidateBookingCache(dateStr);
    logAction(actor || "client", "Отмена", "Booking", bookingId, oldStatus, "❌ Скасовано клієнтом", bookingPhone);
    return { success: true };
  }, actor || "client");
}

function rescheduleBooking(bookingId, phone, newDate, newTime, newBox, actor) {
  return withScriptLock("rescheduleBooking", function() {
    const booking = findBookingById(bookingId);
    if (!booking) return { success: false, error: "booking not found" };
    const bookingPhone = normalizePhone(String(booking.values[7] || ""));
    if (phone && bookingPhone !== normalizePhone(phone)) return { success: false, error: "phone mismatch" };
    if (isBoxBooked(newDate, newTime, newBox)) return { success: false, error: "box busy" };

    const oldDate = booking.values[1] instanceof Date ? fmtDate(booking.values[1]) : String(booking.values[1]).trim();
    const oldValue = `${oldDate} ${booking.values[3]} ${booking.values[4]}`;
    const [day, month, year] = newDate.split(".");
    const d = new Date(parseInt(year), parseInt(month) - 1, parseInt(day));
    const dayName = DAY_NAMES[d.getDay()];
    getSheet(SHEET_BOOKINGS).getRange(booking.row, 2, 1, 4).setValues([[
      newDate, dayName, normalizeTime(newTime), newBox
    ]]);
    getSheet(SHEET_BOOKINGS).getRange(booking.row, 10).setValue("✅ Підтверджено");
    invalidateBookingCache(oldDate);
    invalidateBookingCache(newDate);
    logAction(actor || "client", "Перенос", "Booking", bookingId, oldValue, `${newDate} ${newTime} ${newBox}`, bookingPhone);
    return { success: true };
  }, actor || "client");
}

// ─────────────────────── SUBSCRIPTION REQUESTS ─────────────────

function getSubRequestSheet() {
  let s = getSheet(SHEET_SUB_REQUESTS);
  if (s) return s;
  for (const alt of SHEET_SUB_REQUESTS_ALTS) { s = getSheet(alt); if (s) return s; }
  return null;
}

function saveSubscriptionRequest(clientId, clientName, telegramId, typeId, typeName, amount) {
  return withScriptLock("saveSubscriptionRequest", function() {
    const sheet = getSubRequestSheet();
    if (!sheet) return false;
    assertForeignKeys({ clientId, typeId });
    const row = getNextRow(sheet.getName());
    const requestId = genSubRequestId();
    sheet.getRange(row, 1, 1, 12).setValues([[
      requestId, clientId, clientName, telegramId, typeId, typeName,
      Number(amount) || 0, fmtDate(new Date()), "🆕 Нова заявка", "", "", ""
    ]]);
    sheet.getRange(row, 7).setNumberFormat("#,##0");
    logAction(telegramId || "web", "Заявка на абонемент", "SubscriptionRequest", requestId, "", "🆕 Нова заявка", typeName);
    return requestId;
  }, telegramId || "web");
}

function getSubscriptionRequests() {
  const sheet = getSubRequestSheet();
  if (!sheet) return [];
  return safeGetRangeValues(sheet, 2, 1, null, 12).filter(r => r[0]).map((r, i) => ({
    row: i + 2, requestId: String(r[0]), clientId: String(r[1]), clientName: String(r[2]),
    telegramId: String(r[3]), typeId: String(r[4]), typeName: String(r[5]),
    amount: r[6] || 0, createdDate: parseDate(r[7]), status: String(r[8]),
    rejectReason: String(r[9] || ""), processedDate: parseDate(r[10]), processedBy: String(r[11] || "")
  }));
}

function updateSubscriptionRequestStatus(requestId, newStatus, actor, reason) {
  const sheet = getSubRequestSheet();
  if (!sheet) return false;
  const data = safeGetRangeValues(sheet, 2, 1, null, 12);
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]) === String(requestId)) {
      sheet.getRange(i + 2, 9, 1, 4).setValues([[
        newStatus,
        reason || data[i][9] || "",
        fmtDate(new Date()),
        actor || ""
      ]]);
      logAction(actor || "admin", "Обработка заявки", "SubscriptionRequest", requestId, data[i][8], newStatus, reason || "");
      return true;
    }
  }
  return false;
}

function sellSubscription(clientId, typeId, actor, options, skipLock) {
  const runner = function() {
    assertForeignKeys({ clientId, typeId });
    const tClient = findClientById(clientId);
    const type = getTypes().find(t => String(t.id) === String(typeId));
    if (!type || !tClient) return { success: false, error: "client or type not found" };

    const sheet = getSheet(SHEET_SUBS);
    const row = getNextRow(SHEET_SUBS);
    const subId = genSubId();
    const pDate = new Date();
    const eDate = new Date();
    eDate.setMonth(eDate.getMonth() + (Number(type.months) || 1));
    const amount = Number(options?.amount || type.amount || 0);
    const rowValues = [
      clientId,
      subId,
      typeId,
      options?.clientName || tClient.name,
      options?.typeName || type.name,
      amount,
      fmtDate(pDate),
      fmtDate(eDate),
      0,
      amount,
      "✅ Активний",
      actor || "admin",
      options?.paymentMethod || "",
      options?.discount || "",
      options?.promoCode || "",
      options?.activationDate || fmtDate(pDate)
    ];
    sheet.getRange(row, 1, 1, rowValues.length).setValues([rowValues]);
    const newBalance = updateClientBalance(clientId);
    updateClientMetrics(clientId);
    logAction(actor || "admin", "Продажа абонемента", "Subscription", subId, "", amount, type.name);
    return { success: true, subId, endDate: fmtDate(eDate), amount, newBalance };
  };
  return skipLock ? runner() : withScriptLock("sellSubscription", runner, actor || "admin");
}

function confirmSubscriptionRequest(requestId, actor) {
  return withScriptLock("confirmSubscriptionRequest", function() {
    const req = getSubscriptionRequests().find(r => r.requestId === requestId);
    if (!req) return false;
    assertForeignKeys({ clientId: req.clientId, typeId: req.typeId });
    const sold = sellSubscription(req.clientId, req.typeId, actor || "admin", {
      clientName: req.clientName,
      typeName: req.typeName,
      amount: req.amount,
      paymentMethod: "Заявка",
      activationDate: fmtDate(new Date())
    }, true);
    if (!sold || !sold.success) return false;
    updateSubscriptionRequestStatus(requestId, "✅ Підтверджена", actor || "admin", "");
    return true;
  }, actor || "admin");
}

// ─────────────────────── SUB STATUS HELPER ──────────────────────

/** Змінює статус абонементу (для бота) */
function updateSubStatusBot(subId, newStatus) {
  return withScriptLock("updateSubStatusBot", function() {
    const sheet = getSheet(SHEET_SUBS);
    if (!sheet) return false;
    const data = safeGetRangeValues(sheet, 2, 1, null, 16);
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][1]).trim() === String(subId).trim()) {
        const oldStatus = String(data[i][10] || "");
        sheet.getRange(i + 2, 11).setValue(newStatus);
        try { if (data[i][0]) { updateClientBalance(String(data[i][0])); updateClientMetrics(String(data[i][0])); } } catch(e) {}
        logAction("Telegram admin", "Редактирование", "Subscription", subId, oldStatus, newStatus, "status");
        return true;
      }
    }
    return false;
  }, "Telegram admin");
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
  requireConfig("BOT_TOKEN");
  requireConfig("WEBAPP_URL");
  const url = `https://api.telegram.org/bot${BOT_TOKEN}/setWebhook?url=${WEBAPP_URL}`;
  try { UrlFetchApp.fetch(url); Logger.log("✅ Webhook встановлено"); }
  catch(e) { Logger.log("❌ Webhook: " + e); }
}

function doPost(e) {
  try {
    ensureSchema();
    const update = JSON.parse(e.postData.contents);
    if (update.message) handleMessage(update.message);
    if (update.callback_query) handleCallback(update.callback_query);
  } catch(err) { Logger.log("doPost error: " + err); logError("doPost", err, "telegram"); }
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
    try {
      const clientId = autoRegisterClient(text, state.phone, String(userId), "Telegram");
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

    subs.forEach((s, i) => {
      const msg =
        `📋 <b>${i+1}. ${s.typeName}</b>\n` +
        `🆔 ${s.subId}\n` +
        `💰 Залишок: <b>${fmt(s.balance)} грн</b>\n` +
        `✅ Витрачено: ${fmt(s.spent)} грн\n` +
        `📅 До: ${s.endDate}\n` +
        `${s.status}`;

      // ✅ Кнопки управління залежно від статусу
      const buttons = [];
      if (s.status.includes("Активний")) {
        buttons.push([{ text: "❄️ Заморозити",   callback_data: `sub_freeze_${s.subId}`      }]);
        buttons.push([{ text: "❌ Деактивувати", callback_data: `sub_deactivate_${s.subId}` }]);
      } else if (s.status.includes("Заморожений")) {
        buttons.push([{ text: "✅ Активувати",    callback_data: `sub_activate_${s.subId}`    }]);
        buttons.push([{ text: "❌ Деактивувати", callback_data: `sub_deactivate_${s.subId}` }]);
      } else {
        buttons.push([{ text: "✅ Активувати",    callback_data: `sub_activate_${s.subId}`    }]);
      }
      const keyboard = buttons.length ? createInlineKeyboard(buttons) : null;
      sendMessage(chatId, msg, keyboard);
    });
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

    // ✅ Фільтрація минулих годин для сьогодні
    const todayStr = Utilities.formatDate(new Date(), TZ, "dd.MM.yyyy");
    const isToday  = (dateStr === todayStr);
    let allTimes   = getTimesForDate(dateStr);

    if (isToday) {
      const now     = new Date();
      const curH    = parseInt(Utilities.formatDate(now, TZ, "HH"));
      const curM    = parseInt(Utilities.formatDate(now, TZ, "mm"));
      const minHour = curM > 0 ? curH + 1 : curH;
      allTimes = allTimes.filter(t => parseInt(t.split(':')[0]) >= minHour);
    }

    if (!allTimes.length) {
      sendMessage(chatId, "❌ На цю дату нема доступного часу. Оберіть іншу дату.");
      return;
    }
    setState(userId, { action: "booking", date: dateStr });
    const buttons = allTimes.map(t => [{ text: `⏰ ${t}`, callback_data: `book_time_${t}` }]);
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
    const ok = saveBooking(state.date, state.time, box, client.clientId, client.name, client.phone, "Запис через бот", "Telegram", "", "Telegram " + userId);
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
    const ok = confirmSubscriptionRequest(requestId, "Telegram admin " + userId);
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
    const ok = updateSubscriptionRequestStatus(requestId, "❌ Відхилена", "Telegram admin " + userId, "");
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
      const sold = sellSubscription(clientId, typeId, "Telegram admin " + userId, { paymentMethod: "Telegram" });
      if (!sold?.success) { sendMessage(chatId, "❌ Помилка продажу"); return; }
      sendMessage(chatId, `✅ <b>Продано!</b>\n💎 ${type.name}\n💰 ${sold.amount} грн\n👤 ${tClient.name}\n💳 Новий баланс: ${fmt(sold.newBalance)} грн`);
      
      if (tClient.telegramId && tClient.telegramId !== '0' && tClient.telegramId !== '')
        sendMessage(tClient.telegramId, `🎉 <b>Ви придбали ${type.name}!</b>\n\n💰 ${sold.amount} грн на рахунку\n📅 До: ${sold.endDate}\n\nДякуємо! 🚗`);
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

  // ✅ УПРАВЛІННЯ АБОНЕМЕНТАМИ — заморозка, активація, деактивація
  if (data.startsWith("sub_freeze_")) {
    const subId = data.replace("sub_freeze_", "");
    const ok = updateSubStatusBot(subId, "❄️ Заморожений");
    answerCallback(callbackQuery.id, ok ? "❄️ Заморожено" : "❌ Помилка");
    sendMessage(chatId, ok
      ? `❄️ <b>Абонемент заморожено</b>\n\nID: ${subId}\nДля розморозки натисніть ✅ Активувати`
      : "❌ Не вдалося заморозити абонемент");
    return;
  }

  if (data.startsWith("sub_activate_")) {
    const subId = data.replace("sub_activate_", "");
    const ok = updateSubStatusBot(subId, "✅ Активний");
    answerCallback(callbackQuery.id, ok ? "✅ Активовано" : "❌ Помилка");
    sendMessage(chatId, ok
      ? `✅ <b>Абонемент активовано</b>\n\nID: ${subId}`
      : "❌ Не вдалося активувати абонемент");
    return;
  }

  if (data.startsWith("sub_deactivate_")) {
    const subId = data.replace("sub_deactivate_", "");
    const ok = updateSubStatusBot(subId, "❌ Неактивний");
    answerCallback(callbackQuery.id, ok ? "❌ Деактивовано" : "❌ Помилка");
    sendMessage(chatId, ok
      ? `❌ <b>Абонемент деактивовано</b>\n\nID: ${subId}`
      : "❌ Не вдалося деактивувати абонемент");
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
      const clientId = autoRegisterClient(state.name, state.phone, text === "-" ? "" : text, "Администратор");
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
      const visit = markVisit(state.clientId, state.subId, state.service, state.cost, text === "-" ? "" : text, "Telegram admin " + userId);
      if (!visit?.success) {
        sendMessage(chatId, `❌ Недостатньо коштів на абонементі!\nПеревірте баланс та повторіть.`);
        deleteState(userId);
        return;
      }
      const newBalance = visit.afterBalance;
      
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
  const data = safeGetRangeValues(clientSheet, 2, 1, null, 13);
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
  const rows = calculateAnalytics();
  updateAnalyticsSheet();
  sendMessage(chatId,
    `📊 <b>АНАЛІТИКА</b>\n\n` +
    rows.map(r => `• ${r[0]}: <b>${r[1]}</b>`).join("\n")
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

// ─────────────────────── ADD TO HISTORY ────────────────────────
function addToHistory(bookingId, clientId, clientName, date, service, cost, note, actor, balanceChange, beforeBalance, afterBalance, subId) {
  const sheet = getSheet(SHEET_HISTORY);
  if (!sheet) return false;
  try {
    const row = getNextRow(SHEET_HISTORY);
    sheet.getRange(row, 1, 1, 12).setValues([[
      genHistoryId(),
      date || fmtDate(new Date()),
      String(clientId || ""),
      String(clientName || ""),
      String(subId || ""),
      String(service || ""),
      cost || 0,
      String(note || ""),
      String(actor || ""),
      balanceChange === undefined ? "" : balanceChange,
      beforeBalance === undefined ? "" : beforeBalance,
      afterBalance === undefined ? "" : afterBalance
    ]]);
    Logger.log("✅ Додано до Історії: " + (clientName||clientId) + " — " + service);
    return true;
  } catch(e) {
    Logger.log("❌ addToHistory error: " + e);
    return false;
  }
}

// ─────────────────────── doGet (REST API) ──────────────────────

function doGet(e) {
  const action = e.parameter.action || '';
  let result = {};
  const cache = CacheService.getScriptCache();
  try {
    ensureSchema();

    if (action === 'verifyCrmPin') {
      return jsonResponse(verifyCrmPin(e.parameter.pin || ""));
    }

    try { runDailyMaintenance(false); } catch (maintenanceErr) { logError("runDailyMaintenance", maintenanceErr, "system"); }

    if (requiresAdminAuth(action) && !isValidAdminToken(e.parameter.adminToken || "")) {
      logAction("web", "Unauthorized API", "CRM", action, "", "", "Нет или истек adminToken");
      return jsonResponse({ success: false, unauthorized: true, error: "Unauthorized" });
    }

  // ── 1. Час + бокси разом (один запит замість двох) ─────────────
  if (action === 'getTimesAndBoxes') {
    const dateStr   = e.parameter.date;
    const todayStr  = Utilities.formatDate(new Date(), TZ, "dd.MM.yyyy");
    const isToday   = (dateStr === todayStr);
    const cacheKey  = isToday ? null : 'timesBoxes_' + dateStr;

    if (cacheKey) {
      const cached = cache.get(cacheKey);
      if (cached) { result = JSON.parse(cached); }
    }

    if (!result.times) {
      // ✅ Фільтр минулого часу для сьогодні
      const now    = new Date();
      const curH   = parseInt(Utilities.formatDate(now, TZ, "HH"));
      const curM   = parseInt(Utilities.formatDate(now, TZ, "mm"));
      const minHour = isToday ? (curM > 0 ? curH + 1 : curH) : 0;

      const allTimes = getTimesForDate(dateStr);
      const times = allTimes.map(time => {
        const timeHour = parseInt(time.split(':')[0]);
        if (isToday && timeHour < minHour) {
          return { time, available: false, past: true, freeBoxes: [] };
        }
        const boxes     = getBoxesForTime(dateStr, time);
        const freeBoxes = boxes.filter(b => !isBoxBooked(dateStr, time, b));
        return { time, available: freeBoxes.length > 0, freeBoxes };
      });

      result = { times, date: dateStr };
      // Кешуємо тільки майбутні дати (2 хвилини)
      if (cacheKey) cache.put(cacheKey, JSON.stringify(result), 120);
    }
  }

  // ── 2. Доступний час (legacy — зі старим клієнтом) ─────────────
  else if (action === 'getAvailableTimesForDate') {
    const dateStr  = e.parameter.date;
    const todayStr = Utilities.formatDate(new Date(), TZ, "dd.MM.yyyy");
    const isToday  = (dateStr === todayStr);
    const cacheKey = isToday ? null : 'avail_' + dateStr;

    if (cacheKey) {
      const cached = cache.get(cacheKey);
      if (cached) { result = JSON.parse(cached); }
    }

    if (!result.times) {
      const now     = new Date();
      const curH    = parseInt(Utilities.formatDate(now, TZ, "HH"));
      const curM    = parseInt(Utilities.formatDate(now, TZ, "mm"));
      const minHour = isToday ? (curM > 0 ? curH + 1 : curH) : 0;

      const allTimes = getTimesForDate(dateStr);
      const times = allTimes.map(time => {
        const timeHour = parseInt(time.split(':')[0]);
        if (isToday && timeHour < minHour) {
          return { time, available: false, past: true, freeBoxes: [] };
        }
        const boxes     = getBoxesForTime(dateStr, time);
        const freeBoxes = boxes.filter(b => !isBoxBooked(dateStr, time, b));
        return { time, available: freeBoxes.length > 0, freeBoxes };
      });

      result = { times, date: dateStr };
      if (cacheKey) cache.put(cacheKey, JSON.stringify(result), 120);
    }
  }

  // ── 3. Вільні бокси для часу ────────────────────────────────────
  else if (action === 'getBoxesForDateTime') {
    result = { boxes: getAvailableBoxes(e.parameter.date, e.parameter.time) };
  }

  else if (action === 'cancelBooking') {
    result = cancelBooking(e.parameter.bookingId, e.parameter.phone, "client");
  }

  else if (action === 'rescheduleBooking') {
    result = rescheduleBooking(
      e.parameter.bookingId,
      e.parameter.phone,
      e.parameter.date,
      e.parameter.time,
      e.parameter.box,
      "client"
    );
  }

  // ── 4. Усі записи ───────────────────────────────────────────────
  else if (action === 'getAllBookings') {
    const sheet = getSheet(SHEET_BOOKINGS);
    const data  = safeGetRangeValues(sheet, 2, 1, null, 12);
    result = {
      bookings: data.filter(r => r[0]).map((row, idx) => ({
        row:        idx + 2,
        id:         String(row[0]),
        date:       row[1] instanceof Date ? fmtDate(row[1]) : String(row[1]).trim(),
        day:        String(row[2]),
        time:       row[3] instanceof Date ? Utilities.formatDate(row[3], TZ, "HH:mm") : String(row[3]).trim(),
        box:        String(row[4]),
        clientId:   String(row[5]),
        clientName: String(row[6]),
        phone:      normalizePhone(String(row[7])),
        service:    String(row[8]),
        status:     String(row[9]),
        note:       String(row[10] || ''),
        source:     String(row[11] || '')
      }))
    };
  }

  // ── 5. Усі клієнти ──────────────────────────────────────────────
  else if (action === 'getAllClients') {
    const sheet = getSheet(SHEET_CLIENTS);
    const data  = safeGetRangeValues(sheet, 2, 1, null, 13);
    result = {
      clients: data.filter(r => r[0]).map(row => ({
        clientId:   String(row[0]),
        name:       String(row[1]),
        phone:      normalizePhone(String(row[2])),
        telegramId: String(row[3] || ''),
        status:     String(row[4] || CLIENT_STATUS_ACTIVE),
        registered: parseDate(row[5]),
        balance:    getClientLiveBalance(String(row[0])),
        visits:     Number(row[7]) || 0,
        totalPurchases: Number(row[8]) || 0,
        averageCheck: Number(row[9]) || 0,
        lastVisit:  parseDate(row[10]),
        source:     String(row[11] || ''),
        adminComment: String(row[12] || '')
      }))
    };
  }

  // ── 6. Повна інформація про клієнта ────────────────────────────
  else if (action === 'getClientFull') {
    const cid    = e.parameter.clientId;
    const client = findClientById(cid);
    if (client) client.phone = normalizePhone(String(client.phone || ''));
    result = {
      client,
      subs:     getClientSubs(cid),
      history:  getClientHistory(cid),
      bookings: getClientBookings(cid)
    };
  }

  // ── 7. Усі абонементи БЕЗ візитів (legacy) ─────────────────────
  else if (action === 'getAllSubs') {
    const sheet = getSheet(SHEET_SUBS);
    const data  = safeGetRangeValues(sheet, 2, 1, null, 16);
    result = {
      subs: data.filter(r => r[0]).map((row, idx) => ({
        row: idx + 2, clientId: String(row[0]), subId: String(row[1]),
        typeId: String(row[2]), clientName: String(row[3]), typeName: String(row[4]),
        amount: row[5] || 0, startDate: parseDate(row[6]), endDate: parseDate(row[7]),
        spent: row[8] || 0, balance: row[9] || 0, status: String(row[10]),
        soldBy: String(row[11] || ''), paymentMethod: String(row[12] || ''),
        discount: row[13] || '', promoCode: String(row[14] || ''), activationDate: parseDate(row[15])
      }))
    };
  }

  // ── 8. ✅ Усі абонементи З кількістю візитів ────────────────────
  else if (action === 'getAllSubsWithVisits') {
    // Рахуємо візити з Історії
    const histSheet = getSheet(SHEET_HISTORY);
    const histData  = histSheet ? safeGetRangeValues(histSheet, 2, 1, null, 3) : [];
    const visitMap  = {};
    histData.forEach(r => {
      if (r[0] && r[2]) {
        const cid = String(r[2]).trim();
        visitMap[cid] = (visitMap[cid] || 0) + 1;
      }
    });

    const sheet = getSheet(SHEET_SUBS);
    const data  = safeGetRangeValues(sheet, 2, 1, null, 16);
    result = {
      subs: data.filter(r => r[0]).map((row, idx) => ({
        row:        idx + 2,
        clientId:   String(row[0]),
        subId:      String(row[1]),
        typeId:     String(row[2]),
        clientName: String(row[3]),
        typeName:   String(row[4]),
        amount:     parseFloat(row[5]) || 0,
        startDate:  parseDate(row[6]),
        endDate:    parseDate(row[7]),
        spent:      parseFloat(row[8]) || 0,
        balance:    parseFloat(row[9]) || 0,
        status:     String(row[10]),
        visitCount: visitMap[String(row[0])] || 0,
        soldBy:     String(row[11] || ''),
        paymentMethod: String(row[12] || ''),
        discount:   row[13] || '',
        promoCode:  String(row[14] || ''),
        activationDate: parseDate(row[15])
      }))
    };
  }

  // ── 9. ✅ Оновити статус запису + запис в Історію ───────────────
  else if (action === 'updateBookingStatus') {
    result = withScriptLock("updateBookingStatus", function() {
      const bookingId = e.parameter.bookingId;
      const newStatus = String(e.parameter.status || "");
      const cost = Number(String(e.parameter.cost || 0).replace(/[^\d.]/g, "")) || 0;
      const note = String(e.parameter.note || "Запис виконано");
      const sheet = getSheet(SHEET_BOOKINGS);
      const data = safeGetRangeValues(sheet, 2, 1, null, 12);
      let updated = false;
      let bData = null;
      let oldStatus = "";
      let bookingRow = -1;
      let beforeBalance = 0;
      let subIdForDeduction = null;

      for (let i = 0; i < data.length; i++) {
        if (String(data[i][0]) === bookingId) {
          oldStatus = String(data[i][9] || "");
          bData = {
            date:       data[i][1] instanceof Date ? fmtDate(data[i][1]) : String(data[i][1]).trim(),
            clientId:   String(data[i][5]),
            clientName: String(data[i][6]),
            service:    String(data[i][8])
          };
          bookingRow = i + 2;
          if ((newStatus.includes("Виконано") || newStatus.includes("Готово")) && cost > 0) {
            subIdForDeduction = e.parameter.subId || findSubForCost(bData.clientId, cost);
            if (!subIdForDeduction) return { success: false, error: "Абонемент не знайдено або недостатньо коштів" };
            beforeBalance = getClientLiveBalance(bData.clientId);
            const deducted = deductFromSub(bData.clientId, subIdForDeduction, cost);
            if (!deducted) return { success: false, error: "Недостатньо коштів на абонементі" };
          }
          sheet.getRange(bookingRow, 10).setValue(newStatus);
          updated = true;
          break;
        }
      }

      if (updated && bData && (newStatus.includes('Виконано') || newStatus.includes('Готово'))) {
        if (cost > 0 && subIdForDeduction) {
          const afterBalance = getClientLiveBalance(bData.clientId);
          addToHistory(bookingId, bData.clientId, bData.clientName,
                       bData.date, bData.service, cost, note, 'CRM',
                       afterBalance - beforeBalance, beforeBalance, afterBalance, subIdForDeduction);
        } else {
          addToHistory(bookingId, bData.clientId, bData.clientName,
                       bData.date, bData.service, 0, note, 'CRM', 0, "", "", "");
        }
        updateClientMetrics(bData.clientId);
      }
      if (bData?.date) invalidateBookingCache(bData.date);
      if (updated) logAction("CRM", newStatus.includes("Скасовано") ? "Отмена" : "Редактирование", "Booking", bookingId, oldStatus, newStatus, "");
      return { success: updated };
    }, "CRM");
  }

  // ── 10. Оновити статус абонементу ──────────────────────────────
  else if (action === 'updateSubStatus') {
    result = withScriptLock("updateSubStatus", function() {
      const sheet = getSheet(SHEET_SUBS);
      const data  = safeGetRangeValues(sheet, 2, 1, null, 16);
      let updated = false;
      let clientId = "";
      let oldStatus = "";
      for (let i = 0; i < data.length; i++) {
        if (String(data[i][1]) === e.parameter.subId) {
          clientId = String(data[i][0] || "");
          oldStatus = String(data[i][10] || "");
          sheet.getRange(i + 2, 11).setValue(e.parameter.status);
          updated = true; break;
        }
      }
      if (updated && clientId) {
        updateClientBalance(clientId);
        updateClientMetrics(clientId);
        logAction("CRM", "Редактирование", "Subscription", e.parameter.subId, oldStatus, e.parameter.status, "status");
      }
      return { success: updated };
    }, "CRM");
  }

  // ── 11. ✅ Зберегти запис з сайту + автореєстрація ──────────────
  else if (action === 'saveBooking') {
    const name  = e.parameter.name  || 'Гість';
    const phone = normalizePhone(e.parameter.phone || '');
    const clientId = autoRegisterClient(name, phone, "", "Сайт");
    if (!clientId) {
      result = { success: false, error: 'Не вдалося зареєструвати клієнта' };
    } else {
    const bookingId = saveBooking(
      e.parameter.date, e.parameter.time, e.parameter.box,
      clientId, name, phone,
      (e.parameter.service || '') + (e.parameter.car ? ' | ' + e.parameter.car : ''),
      "Сайт",
      e.parameter.comment || "",
      "web"
    );
    if (bookingId) {
      // Інвалідуємо кеш часів
      invalidateBookingCache(e.parameter.date);
      sendMessage(ADMIN_ID,
        `🌐 <b>ЗАПИС З САЙТУ!</b>\n\n` +
        `👤 ${name}\n📱 ${phone}\n` +
        `📅 ${e.parameter.date} ${e.parameter.time}\n` +
        `🏠 ${e.parameter.box}\n🔧 ${e.parameter.service}\n` +
        `🚗 ${e.parameter.car||'—'} · ${e.parameter.carType||'—'}`
      );
    }
    result = { success: !!bookingId, bookingId: bookingId || null };
    }
  }

  // ── 12. Заявка на абонемент з сайту ────────────────────────────
  else if (action === 'saveSubOrder') {
    try {
      const name     = e.parameter.name  || 'Гість';
      const phone    = normalizePhone(e.parameter.phone || '');
      const clientId = autoRegisterClient(name, phone, e.parameter.tg || '', "Сайт");
      if (!clientId) throw new Error('Не вдалося зареєструвати клієнта');
      const requestId = saveSubscriptionRequest(
        clientId, name, e.parameter.tg || '',
        e.parameter.subId, e.parameter.subName,
        Number(String(e.parameter.account || e.parameter.price || 0).replace(/[^\d.]/g, '')) || 0
      );
      if (requestId) {
        sendMessage(ADMIN_ID,
          `💎 <b>НОВА ЗАЯВКА НА АБОНЕМЕНТ!</b>\n\n` +
          `👤 ${name}\n📱 ${phone}\n💬 Telegram: ${e.parameter.tg||'—'}\n\n` +
          `🏷 <b>${e.parameter.subName}</b>\n📅 ${e.parameter.months} міс.\n` +
          `💰 ${Number(String(e.parameter.price||0).replace(/[^\d.]/g,'')).toLocaleString('uk-UA')} грн` +
          ` → ${Number(String(e.parameter.account||0).replace(/[^\d.]/g,'')).toLocaleString('uk-UA')} грн (+${e.parameter.bonus})\n\n` +
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

  // ── 13. Заявки на абонементи ────────────────────────────────────
  else if (action === 'getAllSubRequests') {
    result = { requests: getSubscriptionRequests() };
  }

  // ── 14. Підтвердити заявку ─────────────────────────────────────
  else if (action === 'confirmSubRequest') {
    const ok = confirmSubscriptionRequest(e.parameter.requestId, "CRM");
    if (ok) {
      const req = getSubscriptionRequests().find(r => r.requestId === e.parameter.requestId);
      if (req?.telegramId && req.telegramId !== 'undefined')
        sendMessage(req.telegramId, `🎉 <b>Заявка підтверджена!</b>\n\n💎 ${req.typeName}\n💰 ${req.amount} грн`);
    }
    result = { success: ok };
  }

  // ── 15. Відхилити заявку ───────────────────────────────────────
  else if (action === 'rejectSubRequest') {
    const ok = updateSubscriptionRequestStatus(e.parameter.requestId, '❌ Відхилена', "CRM", e.parameter.reason || "");
    if (ok) {
      const req = getSubscriptionRequests().find(r => r.requestId === e.parameter.requestId);
      if (req?.telegramId && req.telegramId !== 'undefined')
        sendMessage(req.telegramId, `❌ <b>Заявку відхилено</b>\n\n${req.typeName}\n\n💬 ${PHONE}`);
    }
    result = { success: ok };
  }

  // ── 16. Оновити статус заявки ──────────────────────────────────
  else if (action === 'updateSubRequestStatus') {
    result = { success: updateSubscriptionRequestStatus(e.parameter.requestId, e.parameter.status, "CRM", e.parameter.reason || "") };
  }

  // ── 17. Автореєстрація клієнта ─────────────────────────────────
  else if (action === 'autoRegister') {
    const name  = e.parameter.name  || '';
    const phone = e.parameter.phone || '';
    if (!name || !phone) {
      result = { success: false, error: 'name and phone required' };
    } else {
      const clientId = autoRegisterClient(name, phone, e.parameter.telegramId || "", e.parameter.source || "Сайт");
      result = { success: !!clientId, clientId: clientId || null };
    }
  }

  // ── 18. ✅ Перерахувати всі баланси (кнопка в CRM) ──────────────
  else if (action === 'recalcAllBalances') {
    const count = syncAllBalances();
    const metrics = syncClientMetrics();
    result = { success: true, updated: count, metrics };
  }

  // ── 19. Синхронізація балансів ─────────────────────────────────
  else if (action === 'syncBalances') {
    const count = syncAllBalances();
    const metrics = syncClientMetrics();
    result = { success: true, synced: count, metrics };
  }

  else if (action === 'runMaintenance') {
    result = { success: true, maintenance: runDailyMaintenance(true) };
  }

  else if (action === 'getAnalytics') {
    result = { success: true, analytics: calculateAnalytics().map(r => ({ kpi: r[0], value: r[1] })) };
  }

  else if (action === 'broadcastTelegram') {
    result = broadcastTelegram(e.parameter.message || "", e.parameter.segment || "all", "CRM");
  }

  // ── 20. Баланс одного клієнта ──────────────────────────────────
  else if (action === 'getClientBalance') {
    const cid = e.parameter.clientId;
    result = { clientId: cid, balance: getClientLiveBalance(cid) };
  }

  } catch (err) {
    logError(action || "doGet", err, e?.parameter?.adminToken ? "crm" : "web");
    result = { success: false, error: String(err && err.message ? err.message : err) };
  }

  return jsonResponse(result);
}
