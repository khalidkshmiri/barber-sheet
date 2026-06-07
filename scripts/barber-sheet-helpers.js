/***************
 * HELPERS
 * Small utility functions, index loaders, and date helpers.
 * All functions here are private (trailing underscore) except where noted.
 ***************/

function getStandardServicePrice_(serviceName) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SERVICES_SHEET);
  if (!sheet || !serviceName) return 15;
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_ROW) return 15;
  // Read cols B-C (ServiceName, Price) — 2 cols from col 2
  const data = sheet.getRange(DATA_ROW, 2, lastRow - 2, 2).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase() === String(serviceName).toLowerCase()) {
      let p = data[i][1];
      if (typeof p === "string") p = parseFloat(p.replace(/[€\s]/g, "").replace(",", "."));
      return isNaN(p) ? 0 : p;
    }
  }
  return 15;
}

function loadServicePrices_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_ROW) return {};
  // Read cols B-C (ServiceName, Price)
  const data = sheet.getRange(DATA_ROW, 2, lastRow - 2, 2).getValues();
  const prices = {};
  for (let i = 0; i < data.length; i++) {
    let p = data[i][1];
    if (typeof p === "string") p = parseFloat(p.replace(/[€\s]/g, "").replace(",", "."));
    if (data[i][0]) prices[String(data[i][0]).toLowerCase()] = isNaN(p) ? 0 : p;
  }
  return prices;
}

function loadClientsIndex_(sheet) {
  const map = new Map();
  let maxId = 0;
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_ROW) return { map, maxId };
  // Read 16 cols from col B (B-Q): index 0=Name, index 15=ClientID
  const data = sheet.getRange(DATA_ROW, 2, lastRow - 2, 16).getValues();
  for (let i = 0; i < data.length; i++) {
    const name = nameCase_(data[i][0]);   // index 0 = col B = Name
    const id = data[i][15];               // index 15 = col Q = ClientID
    if (name) map.set(name, id);
    const n = parseInt(id, 10);
    if (!isNaN(n) && n > maxId) maxId = n;
  }
  return { map, maxId };
}

function getOrCreateClientId_(sheet, idx, name) {
  const formatted = nameCase_(name);
  if (idx.map.has(formatted)) return idx.map.get(formatted);

  idx.maxId += 1;
  const newRow = sheet.getLastRow() + 1;
  if (sheet.getMaxColumns() < 19) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), 19 - sheet.getMaxColumns());
  }
  sheet.getRange(newRow, 2).setValue(formatted);   // col B = Name
  sheet.getRange(newRow, 17).setValue(idx.maxId);  // col Q = ClientID
  idx.map.set(formatted, idx.maxId);
  return idx.maxId;
}

function loadActiveSubscriptionsIndex_(sheet) {
  const today = startOfDay_(new Date());
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_ROW) return { byId: new Map(), byName: new Map() };
  // Subscriptions layout from col B: B=Name(0) C=Price(1) D=Type(2) E=Expiry(3) F=Credits(4) G=Status(5) H=Notes(6) I=StartDate(7) J=ClientID(8)
  const data = sheet.getRange(DATA_ROW, 2, lastRow - 2, 9).getValues();
  const byId = new Map(), byName = new Map();
  for (let i = 0; i < data.length; i++) {
    const name     = data[i][0]; // col B = Name
    const credits  = data[i][4]; // col F = Credits
    const status   = data[i][5]; // col G = Status
    const expiryRaw = data[i][3]; // col E = Expiry
    const clientId = data[i][8]; // col J = ClientID
    const expiry = expiryRaw ? new Date(expiryRaw) : null;
    if (status !== "Active" || !expiry || expiry < today) continue;
    if (credits !== "" && Number(credits) <= 0) continue;
    const entry = { start: data[i][7] ? new Date(data[i][7]) : new Date(0) }; // col I = StartDate
    if (clientId) byId.set(String(clientId), entry);
    if (name && !String(name).startsWith("=")) byName.set(nameCase_(name), entry);
  }
  return { byId, byName };
}

function hasActiveCredits_(idx, name, id) {
  return (id && idx.byId.has(String(id))) || (name && idx.byName.has(nameCase_(name)));
}

function loadAppointmentEventIdIndex_(sheet) {
  const map = new Map();
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_ROW) return map;
  // Read 13 cols from col B (B-N): index 12=EventID (col N), index 11=ClientID (col M)
  const data = sheet.getRange(DATA_ROW, 2, lastRow - 2, 13).getValues();
  for (let i = 0; i < data.length; i++) {
    const eventId = data[i][12]; // index 12 = col N = EventID
    if (eventId) map.set(String(eventId), { row: i + DATA_ROW, data: data[i] });
  }
  return map;
}

function parseEventTitle_(t) {
  if (!t) return null;
  const parts = t.split(/\s*-\s*/);
  return parts.length >= 2
    ? { clientName: nameCase_(parts[0]), service: parts[1].trim() }
    : { clientName: nameCase_(parts[0]), service: "Haircut" };
}

function getCalendarOrThrow_() {
  const c = CalendarApp.getCalendarsByName(CALENDAR_NAME);
  if (!c.length) throw new Error(`Calendar "${CALENDAR_NAME}" not found.`);
  return c[0];
}

function getSheetOrThrow_(ss, n) {
  const s = ss.getSheetByName(n);
  if (!s) throw new Error(`Missing sheet: ${n}`);
  return s;
}

function notify_(m, duration) {
  Logger.log(m);
  try { SpreadsheetApp.getActive().toast(m, "✂️ Barber Sheet", duration || 5); } catch (_) {}
}

function sendTelegramError_(message) {
  try {
    const props = PropertiesService.getScriptProperties();
    const token = props.getProperty("TELEGRAM_BOT_TOKEN");
    const chatId = props.getProperty("TELEGRAM_CHAT_ID");
    if (!token || !chatId) return;
    UrlFetchApp.fetch("https://api.telegram.org/bot" + token + "/sendMessage", {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({ chat_id: chatId, text: "⚠️ Barber Sheet Error\n\n" + message }),
      muteHttpExceptions: true
    });
  } catch (_) {}
}

function nameCase_(v) {
  return String(v ?? "").trim().toLowerCase()
    .split(" ").map(w => w ? w[0].toUpperCase() + w.slice(1) : "").join(" ");
}

function toSheetDateTime_(s, tz, isAllDay) {
  const ymd = Utilities.formatDate(s, tz, "yyyy-MM-dd");
  const hm = Utilities.formatDate(s, tz, "HH:mm");
  const [y, m, d] = ymd.split("-").map(Number);
  const timeCell = isAllDay ? "" : (() => {
    const [hh, mm] = hm.split(":").map(Number);
    return new Date(1899, 11, 30, hh, mm);
  })();
  return { dateCell: new Date(y, m - 1, d), timeCell, ymd, hm };
}

function startOfDay_(d) { const x = new Date(d); x.setHours(0, 0, 0, 0); return x; }
function endOfDay_(d) { const x = new Date(d); x.setHours(23, 59, 59, 999); return x; }
function addDays_(d, n) { const x = new Date(d); x.setDate(x.getDate() + n); return x; }
function ymd_(d, tz) { return Utilities.formatDate(d, tz, "yyyy-MM-dd"); }
function hm_(d, tz) { return Utilities.formatDate(d, tz, "HH:mm"); }
