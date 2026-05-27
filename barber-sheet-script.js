/***************
 * CONFIGURATION
 ***************/
const CALENDAR_NAME = "Barber Appointments";
const APPOINTMENTS_SHEET = "Appointments";
const CLIENTS_SHEET = "Clients";
const SERVICES_SHEET = "Services";
const SUBSCRIPTIONS_SHEET = "Subscriptions";

const DAYS_BACK = 14;
const DAYS_FORWARD = 14;
const HIDE_OLDER_THAN_DAYS = 30;

// ─── NOTIFICATION MODE ───────────────────────────────────────────────
const NOTIFICATION_MODE = "telegram";
// ─────────────────────────────────────────────────────────────────────

// Appointments columns (1-based):
// A=1 Date  B=2 Time  C=3 Name(formula)  D=4 Price  E=5 Payment
// F=6 Status  G=7 Tips  H=8 Late  I=9 Notes  J=10 Service
// K=11 ClientID  L=12 EventID  M=13 CachedName

// Clients columns (1-based):
// A=1 Name  B=2 FavService  C=3 LastVisit  D=4 SocialMedia  E=5 Notes
// F=6 NoShow(12m)  G=7 Late(12m)  H=8 Referral  I=9 TotalVisits
// J=10 TotalTips  K=11 TotalSpent  L=12 ClientID  M=13 FirstVisit
// N=14 DoNotCut  O=15 ConsecutivePaid (auto-updated)  P=16 VIP (checkbox)

/***************
 * MENU & TRIGGERS
 ***************/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("✂️ Barber Tools")
    .addItem("🔄 Sync Now", "syncCalendarToSheets")
    .addItem("⚡ Setup Incremental Sync (5 min)", "setupIncrementalSync")
    .addItem("🗑️ Stop Incremental Sync", "removeIncrementalSync")
    .addToUi();
}

/**
 * INSTALLABLE TRIGGER — linked to 'On edit' in Script Settings
 */
function processSheetChanges(e) {
  const range = e && e.range;
  if (!range) return;
  const sheet = range.getSheet();
  const sheetName = sheet.getName();

  // 1. Dashboard mobile sync button (C2)
  if (sheetName === "Dashboard" && range.getRow() === 2 && range.getColumn() === 3) {
    if (range.getValue() === true) {
      range.setValue(false);
      SpreadsheetApp.flush();
      syncCalendarToSheets();
    }
    return;
  }

  // 2. Appointments sheet logic
  if (sheetName === APPOINTMENTS_SHEET && range.getRow() > 1) {
    const row = range.getRow();

    // A. Payment type changed (col E)
    if (range.getColumn() === 5) {
      const paymentVal = range.getValue();
      const statusRange = sheet.getRange(row, 6);
      const currentStatus = statusRange.getValue();

      // Never override a manually set No Show or Cancelled
      if (currentStatus !== "No Show" && currentStatus !== "Cancelled") {
        if (paymentVal === "Cash" || paymentVal === "Tikkie" || paymentVal === "Subscription") {
          statusRange.setValue("Paid");
        } else if (paymentVal === "" && currentStatus === "Paid") {
          const dateVal = sheet.getRange(row, 1).getValue();
          const isUpcoming = dateVal && new Date(dateVal) >= startOfDay_(new Date());
          statusRange.setValue(isUpcoming ? "Upcoming" : "Not Paid");
        }
      }

      // Price automation
      const priceRange = sheet.getRange(row, 4);
      const currentPrice = priceRange.getValue();
      const serviceName = String(sheet.getRange(row, 10).getValue()).toLowerCase();

      if (paymentVal === "Subscription") {
        priceRange.setValue(0);
      } else if (paymentVal === "" && currentPrice === 0) {
        priceRange.setValue(getStandardServicePrice_(serviceName));
      }
    }

    // B. Status changed (col F)
    if (range.getColumn() === 6) {
      const statusVal = range.getValue();
      
      // Clear Late checkbox ONLY for No Show / Cancelled
      if (statusVal === "No Show" || statusVal === "Cancelled") {
        sheet.getRange(row, 8).setValue(false);
      }
      
      // If setting to Free, ensure Price = 0 (but keep Late checkbox)
      if (statusVal.startsWith("Free")) {
        sheet.getRange(row, 4).setValue(0);
      }
    }
  }

  // 3. Client name auto-formatting (col A)
  if (sheetName === CLIENTS_SHEET && range.getColumn() === 1 && range.getRow() > 1) {
    const v = range.getValue();
    const fixed = nameCase_(v);
    if (fixed !== v) range.setValue(fixed);
  }
}

/***************
 * MAIN SYNC ENGINES
 ***************/
function syncCalendarToSheets(showNotification = true) {
  const ctx = prepareContext_();
  const now = new Date();
  const startDate = startOfDay_(addDays_(now, -DAYS_BACK));
  const endDate = endOfDay_(addDays_(now, DAYS_FORWARD));
  const events = ctx.calendar.getEvents(startDate, endDate);

  const counts = upsertEvents_(events, { ...ctx, startDate, endDate });
  updateUpcomingToNotPaid_(ctx.appointmentsSheet);
  updateConsecutivePaidCounts_(ctx);
  sortAndHideAppointments_(ctx.appointmentsSheet);
  
  // Only show alert and send notification if it's a manual sync (showNotification = true)
  if (showNotification) {
    safeAlert_(`Sync Complete!\n\n+ ${counts.newCount} New\n~ ${counts.updatedCount} Updated\n- ${counts.cancelledCount} Cancelled`);
    
    // Send sync notification ONLY if there were actual changes
    if (counts.newCount > 0 || counts.cancelledCount > 0) {
      sendSyncNotification_(ctx, counts.newEventIds);
    }
  }
}

function syncThisYear(showNotification = true) {
  const ctx = prepareContext_();
  const now = new Date();
  const year = now.getFullYear();
  const startDate = new Date(year, 0, 1);
  const endDate = new Date(year, 11, 31, 23, 59, 59, 999);

  const events = ctx.calendar.getEvents(startDate, endDate);
  const counts = upsertEvents_(events, { ...ctx, startDate, endDate });
  updateUpcomingToNotPaid_(ctx.appointmentsSheet);
  updateConsecutivePaidCounts_(ctx);
  sortAndHideAppointments_(ctx.appointmentsSheet);
  
  // Only show alert and send notification if it's a manual sync (showNotification = true)
  if (showNotification) {
    safeAlert_(`Sync Complete!\n\n+ ${counts.newCount} New\n~ ${counts.updatedCount} Updated\n- ${counts.cancelledCount} Cancelled`);
    
    // Send sync notification ONLY if there were actual changes
    if (counts.newCount > 0 || counts.cancelledCount > 0) {
      sendSyncNotification_(ctx, counts.newEventIds);
    }
  }
}

function prepareContext_() {
  const ss = SpreadsheetApp.getActive();
  const cal = getCalendarOrThrow_();
  return {
    ss,
    appointmentsSheet: getSheetOrThrow_(ss, APPOINTMENTS_SHEET),
    clientsSheet: getSheetOrThrow_(ss, CLIENTS_SHEET),
    servicesSheet: getSheetOrThrow_(ss, SERVICES_SHEET),
    subscriptionsSheet: getSheetOrThrow_(ss, SUBSCRIPTIONS_SHEET),
    calendar: cal,
    calTz: cal.getTimeZone(),
    servicePrices: loadServicePrices_(getSheetOrThrow_(ss, SERVICES_SHEET)),
    clientsIndex: loadClientsIndex_(getSheetOrThrow_(ss, CLIENTS_SHEET)),
    subsIndex: loadActiveSubscriptionsIndex_(getSheetOrThrow_(ss, SUBSCRIPTIONS_SHEET)),
    apptIndex: loadAppointmentEventIdIndex_(getSheetOrThrow_(ss, APPOINTMENTS_SHEET))
  };
}

function upsertEvents_(events, ctx) {
  const today = startOfDay_(new Date());
  let newCount = 0, updatedCount = 0, cancelledCount = 0;
  const validEventIds = new Set();
  const newEventIds = new Set(); // track IDs of newly added events
  const newRowsABC = [], newRowsDToM = [];

  for (const event of events) {
    const eventId = String(event.getId());
    validEventIds.add(eventId);
    const parsed = parseEventTitle_(event.getTitle());
    if (!parsed) continue;

    const existingRow = ctx.apptIndex.get(eventId);
    const clientId = getOrCreateClientId_(ctx.clientsSheet, ctx.clientsIndex, parsed.clientName);
    const hasCredits = hasActiveCredits_(ctx.subsIndex, parsed.clientName, clientId);

    const start = event.getStartTime();
    const { dateCell, timeCell, ymd, hm } = toSheetDateTime_(start, ctx.calTz, event.isAllDayEvent());

    const serviceLower = parsed.service.toLowerCase();
    const isSubscriptionSale = serviceLower === "monthly subscription";

    let price, payment, serviceToWrite = parsed.service;

    if (isSubscriptionSale) {
      serviceToWrite = "Haircut";
      price = 0;
      payment = "Subscription";
    } else if (hasCredits) {
      price = 0;
      payment = "Subscription";
    } else {
      price = ctx.servicePrices[serviceLower] ?? 0;
      payment = "";
    }

    const isFuture = new Date(dateCell) >= today;
    const initialStatus = (payment === "Subscription") ? "Paid" : (isFuture ? "Upcoming" : "Not Paid");

    if (!existingRow) {
      // --- NEW APPOINTMENT ---
      newRowsABC.push([dateCell, timeCell, ""]);
      newRowsDToM.push([
        price, payment, initialStatus, "", false,
        event.getDescription() || "", serviceToWrite, clientId, eventId,
        parsed.clientName
      ]);

      if (isSubscriptionSale) {
        createSubscriptionEntry_(ctx, parsed.clientName, clientId, dateCell);
      }

      ctx.apptIndex.set(eventId, -1);
      newEventIds.add(eventId);
      newCount++;

    } else {
      // --- EXISTING APPOINTMENT ---
      if (existingRow === -1) continue;

      const rowRange = ctx.appointmentsSheet.getRange(existingRow, 1, 1, 13);
      const rowVals = rowRange.getValues()[0];

      const oldDateVal = new Date(rowVals[0]);
      const oldStatus = rowVals[5];
      const oldPayment = String(rowVals[4] || "");
      const oldTimeStr = rowVals[1] ? hm_(new Date(rowVals[1]), ctx.calTz) : "";
      const oldService = String(rowVals[9] || "");
      const oldName = String(rowVals[2] || "");

      // Always flip Upcoming → Not Paid if date has passed
      if (oldStatus === "Upcoming" && oldDateVal < today) {
        ctx.appointmentsSheet.getRange(existingRow, 6).setValue("Not Paid");
      }

      // SEAMLESS SUBSCRIPTION CONVERSION
      if (isSubscriptionSale && oldPayment !== "Subscription") {
        ctx.appointmentsSheet.getRange(existingRow, 4).setValue(0);
        ctx.appointmentsSheet.getRange(existingRow, 5).setValue("Subscription");
        ctx.appointmentsSheet.getRange(existingRow, 6).setValue("Paid");
        createSubscriptionEntry_(ctx, parsed.clientName, clientId, dateCell);
      }

      const changed =
        ymd !== ymd_(oldDateVal, ctx.calTz) ||
        hm !== oldTimeStr ||
        (event.getDescription() || "") !== String(rowVals[8] || "") ||
        serviceToWrite !== oldService ||
        parsed.clientName !== oldName;

      if (changed) {
        let rowPrice = rowVals[3];
        const isUnpaid = oldStatus === "Not Paid" || oldStatus === "Upcoming";
        if (isUnpaid && new Date(dateCell) >= today) {
          rowPrice = (isSubscriptionSale || hasCredits) ? 0 : (ctx.servicePrices[serviceLower] ?? rowPrice);
        }

        const nameFormula = `=XLOOKUP(K${existingRow}; Clients!L:L; Clients!A:A; M${existingRow})`;
        ctx.appointmentsSheet.getRange(existingRow, 1, 1, 3).setValues([[dateCell, timeCell, nameFormula]]);
        ctx.appointmentsSheet.getRange(existingRow, 4).setValue(isSubscriptionSale ? 0 : rowPrice);
        ctx.appointmentsSheet.getRange(existingRow, 9, 1, 5).setValues([[
          event.getDescription() || "", serviceToWrite, clientId, eventId, parsed.clientName
        ]]);
        updatedCount++;
      }
    }
  }

  // Mark cancelled appointments (calendar events deleted)
  ctx.apptIndex.forEach((rowIdx, eId) => {
    if (rowIdx === -1) return;
    if (validEventIds.has(eId)) return;

    const dateVal = ctx.appointmentsSheet.getRange(rowIdx, 1).getValue();
    if (!dateVal) return;
    const rowDate = new Date(dateVal);
    if (rowDate < ctx.startDate || rowDate > ctx.endDate) return;

    const statusRange = ctx.appointmentsSheet.getRange(rowIdx, 6);
    const currentStatus = statusRange.getValue();
    if (currentStatus === "Paid" || currentStatus === "No Show" || currentStatus === "Cancelled") return;

    statusRange.setValue("Cancelled");
    ctx.appointmentsSheet.getRange(rowIdx, 8).setValue(false);
    cancelledCount++;
  });

  // Write all new rows
  if (newRowsABC.length > 0) {
    const startRow = ctx.appointmentsSheet.getLastRow() + 1;
    for (let i = 0; i < newRowsABC.length; i++) {
      const targetRow = startRow + i;
      newRowsABC[i][2] = `=XLOOKUP(K${targetRow}; Clients!L:L; Clients!A:A; M${targetRow})`;
    }
    ctx.appointmentsSheet.getRange(startRow, 1, newRowsABC.length, 3).setValues(newRowsABC);
    ctx.appointmentsSheet.getRange(startRow, 4, newRowsDToM.length, 10).setValues(newRowsDToM);
  }

  return { newCount, updatedCount, cancelledCount, newEventIds };
}

/**
 * Flip any leftover "Upcoming" appointments whose date has passed to "Not Paid"
 */
function updateUpcomingToNotPaid_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const today = startOfDay_(new Date());
  const data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  for (let i = 0; i < data.length; i++) {
    const dateVal = data[i][0];
    const status = data[i][5];
    if (!dateVal || status !== "Upcoming") continue;
    if (new Date(dateVal) < today) {
      sheet.getRange(i + 2, 6).setValue("Not Paid");
    }
  }
}

/**
 * Update Consecutive Paid counts in Clients sheet (col O)
 */
function updateConsecutivePaidCounts_(ctx) {
  const clientsSheet = ctx.clientsSheet;
  const apptSheet = ctx.appointmentsSheet;
  const lastRow = clientsSheet.getLastRow();
  if (lastRow < 2) return;

  const clientData = clientsSheet.getRange(2, 1, lastRow - 1, 16).getValues();
  const apptData = apptSheet.getRange(2, 1, apptSheet.getLastRow() - 1, 13).getValues();

  for (let i = 0; i < clientData.length; i++) {
    const clientId = clientData[i][11]; // col L
    if (!clientId) continue;

    // Get all appointments for this client, sorted by date desc
    const clientAppts = [];
    for (const row of apptData) {
      if (String(row[10]) === String(clientId)) { // col K = ClientID
        clientAppts.push({
          date: row[0],
          status: row[5],
          payment: row[4],
          late: row[7]  // col H = Late checkbox
        });
      }
    }

    clientAppts.sort((a, b) => new Date(b.date) - new Date(a.date));

    // Count consecutive paid from most recent
    // RESET if: No Show, Late checkbox checked, or unpaid
    let consecutivePaid = 0;
    for (const appt of clientAppts) {
      // Stop counter if No Show or marked Late
      if (appt.status === "No Show" || appt.late === true) {
        break;
      }
      
      const isPaid = appt.payment === "Cash" || appt.payment === "Tikkie" || appt.payment === "Subscription" || appt.status.startsWith("Free");
      if (isPaid) consecutivePaid++;
      else break;
    }

    clientsSheet.getRange(i + 2, 15).setValue(consecutivePaid); // col O
  }
}

function sortAndHideAppointments_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).sort([
    { column: 1, ascending: false },
    { column: 2, ascending: false }
  ]);

  const dates = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  const today = new Date();
  const cutoff = new Date(today);
  cutoff.setDate(today.getDate() - HIDE_OLDER_THAN_DAYS);

  let hideStartRow = -1;
  for (let i = 0; i < dates.length; i++) {
    const raw = dates[i][0];
    if (!raw) continue;
    const d = new Date(raw);
    if (isNaN(d.getTime())) continue;
    if (d < cutoff) { hideStartRow = i + 2; break; }
  }

  sheet.showRows(2, lastRow - 1);
  if (hideStartRow !== -1) {
    const rowsToHide = (lastRow + 1) - hideStartRow;
    if (rowsToHide > 0) sheet.hideRows(hideStartRow, rowsToHide);
  }
}

function createSubscriptionEntry_(ctx, clientName, clientId, startDate) {
  const subsSheet = ctx.subscriptionsSheet;
  const monthlyPrice = ctx.servicePrices["monthly subscription"] ?? 40;
  const name = nameCase_(clientName);

  if (ctx.subsIndex.byName.has(name) || (clientId && ctx.subsIndex.byId.has(String(clientId)))) return;

  const lastRow = subsSheet.getLastRow();
  if (lastRow > 1) {
    const checkRange = subsSheet.getRange(Math.max(2, lastRow - 20), 1, Math.min(21, lastRow - 1), 4).getValues();
    const startYMD = ymd_(startDate instanceof Date ? startDate : new Date(startDate), ctx.calTz);
    for (const r of checkRange) {
      if (nameCase_(r[0]) === name && r[3] && ymd_(new Date(r[3]), ctx.calTz) === startYMD) return;
    }
  }

  const r = lastRow + 1;
  const nameFormula = `=XLOOKUP(I${r}; Clients!L:L; Clients!A:A; "")`;
  const creditsFormula = `=MAX(0; 4 - COUNTIFS(Appointments!$K:$K; I${r}; Appointments!$E:$E; "Subscription"; Appointments!$F:$F; "Paid"; Appointments!$A:$A; ">="&D${r}; Appointments!$A:$A; "<="&(G${r} + 21)))`;

  subsSheet.appendRow([nameFormula, monthlyPrice, "", startDate, creditsFormula, "Active", addDays_(startDate, 31), "", clientId]);

  const newLastRow = subsSheet.getLastRow();
  if (newLastRow >= 2) subsSheet.getRange(2, 1, newLastRow - 1, 9).sort({ column: 4, ascending: false });

  const entry = { start: startDate };
  ctx.subsIndex.byName.set(name, entry);
  if (clientId) ctx.subsIndex.byId.set(String(clientId), entry);
}

/***************
 * ONE-TIME MIGRATION
 ***************/
function runMigration() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(APPOINTMENTS_SHEET);
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) { safeAlert_("No appointment data to migrate."); return; }

  // Ensure col M header exists
  if (!sheet.getRange(1, 13).getValue()) {
    sheet.getRange(1, 13).setValue("Cached Name");
  }

  const data = sheet.getRange(2, 1, lastRow - 1, 13).getValues();
  const cachedNames = [];
  const updatedFormulas = [];

  for (let i = 0; i < data.length; i++) {
    const row = i + 2;
    const currentName = String(data[i][2] || "");
    const existingCache = String(data[i][12] || "");

    cachedNames.push([existingCache || currentName]);
    updatedFormulas.push([`=XLOOKUP(K${row}; Clients!L:L; Clients!A:A; M${row})`]);
  }

  sheet.getRange(2, 13, cachedNames.length, 1).setValues(cachedNames);
  sheet.getRange(2, 3, updatedFormulas.length, 1).setFormulas(updatedFormulas);

  safeAlert_(`✅ Migration complete!\n${cachedNames.length} rows updated.\n\nYou can now hide column M (right-click column M header → Hide column).`);
}

/***************
 * NOTIFICATION SYSTEM
 ***************/
function setupNotificationTrigger() {
  removeNotificationTrigger();
  ScriptApp.newTrigger("sendDailyNotification")
    .timeBased()
    .everyDays(1)
    .atHour(21)
    .create();
  safeAlert_("✅ Daily notification set for 9 PM (21:00) every evening.");
}

function removeNotificationTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "sendDailyNotification") ScriptApp.deleteTrigger(t);
  });
}

function sendDailyNotification() {
  sendTelegramNotification_();
}

/**
 * SYNC NOTIFICATION — sent immediately after sync
 */
function sendSyncNotification_(ctx, newEventIds) {
  const props  = PropertiesService.getScriptProperties();
  const token  = props.getProperty("TELEGRAM_BOT_TOKEN");
  const chatId = props.getProperty("TELEGRAM_CHAT_ID");
  if (!token || !chatId) return;

  const apptSheet    = ctx.appointmentsSheet;
  const clientsSheet = ctx.clientsSheet;
  const tz           = ctx.calTz;
  const lastRow      = apptSheet.getLastRow();
  if (lastRow < 2) return;

  const data      = apptSheet.getRange(2, 1, lastRow - 1, 13).getValues();
  const clientMap = loadClientNotificationMap_(clientsSheet);
  const newAppts  = [];

  for (const row of data) {
    const eventId = String(row[11] || "");
    if (!newEventIds || !newEventIds.has(eventId)) continue;

    const name    = row[2] || row[12];
    const dateVal = row[0];
    const timeVal = row[1];
    const price   = row[3];
    const payment = row[4];
    const notes   = row[8];
    if (!name || !dateVal) continue;

    const dateStr  = Utilities.formatDate(new Date(dateVal), tz, "EEEE d MMMM");
    const timeStr  = timeVal ? Utilities.formatDate(new Date(timeVal), tz, "HH:mm") : "\u2014";
    const priceStr = payment === "Subscription" ? "Sub" : "\u20ac" + price;

    const { badge, doNotCut } = getReliabilityInfo_(name, clientMap);
    const vipBadge            = isVIP_(name, clientsSheet)           ? " \u2b50" : "";
    const consecutivePaid     = getConsecutivePaidCount_(name, clientsSheet);
    const loyaltyLabel        = consecutivePaid >= 5                 ? " \u2705 ELIGIBLE FOR FREE" : "";
    const lastAppt            = getClientLastAppointment_(name, apptSheet);
    const nameDisplay         = doNotCut ? "<u>" + name + "</u>" : name;

    newAppts.push({ name: nameDisplay, dateStr, timeStr, priceStr, badge, vipBadge, loyaltyLabel, lastAppt, notes });
  }

  if (newAppts.length === 0) return;

  const plural = newAppts.length > 1 ? "s" : "";
  let lines = ["\u2702\ufe0f <b>New Appointment" + plural + " Added</b>"];

  for (const a of newAppts) {
    lines.push("");
    lines.push("<b>" + a.dateStr + " at " + a.timeStr + "</b>");
    lines.push("\ud83d\udc64 " + a.name + a.vipBadge + " " + a.badge + a.loyaltyLabel);
    lines.push("\ud83d\udcb6 " + a.priceStr);
    if (a.lastAppt) {
      const notesPart = a.lastAppt.notes ? " \u2014 \"" + a.lastAppt.notes + "\"" : "";
      lines.push("\ud83d\udccb Last: " + a.lastAppt.label + notesPart);
    } else {
      lines.push("\ud83d\udccb First visit");
    }
    if (a.notes) lines.push("\ud83d\udcdd \"" + a.notes + "\"");
  }

  const msg = lines.join("\n");

  UrlFetchApp.fetch("https://api.telegram.org/bot" + token + "/sendMessage", {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({ chat_id: chatId, text: msg, parse_mode: "HTML" })
  });
}
/**
 * Gets tomorrow's appointments with last appointment info
 */
function getTomorrowAppointments_() {
  const ss = SpreadsheetApp.getActive();
  const apptSheet = getSheetOrThrow_(ss, APPOINTMENTS_SHEET);
  const clientsSheet = getSheetOrThrow_(ss, CLIENTS_SHEET);

  const tomorrow = startOfDay_(addDays_(new Date(), 1));
  const tomorrowEnd = endOfDay_(addDays_(new Date(), 1));
  const lastRow = apptSheet.getLastRow();
  if (lastRow < 2) return { appts: [], clientMap: new Map(), tomorrow };

  const data = apptSheet.getRange(2, 1, lastRow - 1, 13).getValues();
  const clientMap = loadClientNotificationMap_(clientsSheet);

  const appts = [];
  for (const row of data) {
    const dateVal = row[0];
    if (!dateVal) continue;
    const d = new Date(dateVal);
    if (d < tomorrow || d > tomorrowEnd) continue;
    const status = row[5];
    if (status === "Cancelled" || status === "No Show") continue;

    const name = row[2] || row[12];
    const lastAppt = getClientLastAppointment_(name, apptSheet);

    appts.push({
      time: row[1],
      name,
      service: row[9],
      price: row[3],
      payment: row[4],
      status,
      clientId: row[10],
      lastAppt
    });
  }

  appts.sort((a, b) => {
    const ta = a.time ? new Date(a.time).getTime() : 0;
    const tb = b.time ? new Date(b.time).getTime() : 0;
    return ta - tb;
  });

  return { appts, clientMap, tomorrow };
}

/**
 * Gets unpaid appointments from today
 */
function getTodayUnpaid_() {
  const ss = SpreadsheetApp.getActive();
  const apptSheet = getSheetOrThrow_(ss, APPOINTMENTS_SHEET);
  const today = startOfDay_(new Date());
  const todayEnd = endOfDay_(new Date());
  const lastRow = apptSheet.getLastRow();
  if (lastRow < 2) return [];

  const data = apptSheet.getRange(2, 1, lastRow - 1, 13).getValues();
  const unpaid = [];

  for (const row of data) {
    const dateVal = row[0];
    if (!dateVal) continue;
    const d = new Date(dateVal);
    if (d < today || d > todayEnd) continue;
    const status = row[5];
    if (status !== "Not Paid") continue;
    unpaid.push({
      time: row[1],
      name: row[2] || row[12],
      service: row[9],
      price: row[3]
    });
  }

  unpaid.sort((a, b) => {
    const ta = a.time ? new Date(a.time).getTime() : 0;
    const tb = b.time ? new Date(b.time).getTime() : 0;
    return ta - tb;
  });

  return unpaid;
}

/**
 * Gets unreliable clients with appointments 2 days AND 1 day ahead
 * DEDUP: exclude if already in tomorrow's section
 */
function getUnreliableAppointments_(tomorrowNames) {
  const ss = SpreadsheetApp.getActive();
  const apptSheet = getSheetOrThrow_(ss, APPOINTMENTS_SHEET);
  const clientsSheet = getSheetOrThrow_(ss, CLIENTS_SHEET);

  const twoDays = startOfDay_(addDays_(new Date(), 2));
  const twoDaysEnd = endOfDay_(addDays_(new Date(), 2));
  const oneDay = startOfDay_(addDays_(new Date(), 1));
  const oneDayEnd = endOfDay_(addDays_(new Date(), 1));

  const lastRow = apptSheet.getLastRow();
  if (lastRow < 2) return [];

  const data = apptSheet.getRange(2, 1, lastRow - 1, 13).getValues();
  const clientMap = loadClientNotificationMap_(clientsSheet);
  const flagged = [];
  const seenNames = new Set(tomorrowNames);

  for (const row of data) {
    const dateVal = row[0];
    if (!dateVal) continue;
    const d = new Date(dateVal);

    const isTwoDays = d >= twoDays && d <= twoDaysEnd;
    const isOneDay = d >= oneDay && d <= oneDayEnd;
    if (!isTwoDays && !isOneDay) continue;

    const status = row[5];
    if (status === "Cancelled" || status === "No Show") continue;

    const name = row[2] || row[12];
    
    // DEDUP: skip if already in tomorrow's section
    if (seenNames.has(nameCase_(name))) continue;

    const info = clientMap.get(nameCase_(name)) || {};
    const noShows = info.noShow || 0;
    const lates = info.late || 0;
    const doNotCut = info.doNotCut || false;
    const total = noShows + lates;

    if (total < 3 && !doNotCut) continue;

    flagged.push({
      time: row[1],
      name,
      service: row[9],
      price: row[3],
      payment: row[4],
      noShows,
      lates,
      doNotCut,
      notes: info.notes || "",
      daysAhead: isTwoDays ? 2 : 1
    });
  }

  return flagged;
}

/**
 * Get previous appointment info for a client
 */
function getClientLastAppointment_(clientName, apptSheet) {
  const today = startOfDay_(new Date());
  const data = apptSheet.getRange(2, 1, apptSheet.getLastRow() - 1, 13).getValues();
  const clientAppts = [];

  for (const row of data) {
    const dateVal = row[0];
    if (!dateVal) continue;
    const d = new Date(dateVal);
    if (d >= today) continue; // Only past appointments
    const name = row[2] || row[12];
    if (nameCase_(name) !== nameCase_(clientName)) continue;

    clientAppts.push({
      date: d,
      status: row[5],
      price: row[3],
      payment: row[4],
      notes: row[8]
    });
  }

  if (clientAppts.length === 0) return null;

  // Get most recent
  clientAppts.sort((a, b) => b.date - a.date);
  const last = clientAppts[0];

  let label = last.status;
  if (last.status === "Paid") {
    label = `Paid (€${last.price})`;
  } else if (last.status.startsWith("Free")) {
    label = last.status; // "Free - Family" etc
  } else if (last.status === "Late") {
    label = "Late";
  }

  return {
    status: last.status,
    label,
    notes: last.notes,
    payment: last.payment,
    price: last.price
  };
}

function getReliabilityInfo_(name, clientMap) {
  const info = clientMap.get(nameCase_(name)) || {};
  const noShows = info.noShow || 0;
  const lates = info.late || 0;
  const doNotCut = info.doNotCut || false;
  const notes = info.notes || "";
  const total = noShows + lates;

  let badge, badgeColor;
  if (doNotCut) {
    badge = "⛔ DO NOT CUT";
    badgeColor = "#8B0000";
  } else if (total >= 3) {
    badge = `⚠️ Unreliable (${noShows} NS, ${lates} late)`;
    badgeColor = "#cc0000";
  } else if (total >= 1) {
    badge = `🟡 Watch (${noShows} NS, ${lates} late)`;
    badgeColor = "#856404";
  } else {
    badge = "✅ Reliable";
    badgeColor = "#155724";
  }

  return { badge, badgeColor, doNotCut, notes };
}

/**
 * Check if client is VIP (col P in Clients sheet)
 */
function isVIP_(clientName, clientsSheet) {
  const data = clientsSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (nameCase_(data[i][0]) === nameCase_(clientName)) {
      return data[i][15] === true; // col P
    }
  }
  return false;
}

/**
 * Get consecutive paid count for client
 */
function getConsecutivePaidCount_(clientName, clientsSheet) {
  const data = clientsSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (nameCase_(data[i][0]) === nameCase_(clientName)) {
      return Number(data[i][14]) || 0; // col O
    }
  }
  return 0;
}

/**
 * TELEGRAM notification with all sections
 */
function sendTelegramNotification_() {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty("TELEGRAM_BOT_TOKEN");
  const chatId = props.getProperty("TELEGRAM_CHAT_ID");

  if (!token || !chatId) {
    safeAlert_("⚠️ Telegram not configured.\n\nGo to: Extensions → Apps Script → Project Settings → Script Properties");
    return;
  }

  const cal = getCalendarOrThrow_();
  const tz = cal.getTimeZone();
  const ss = SpreadsheetApp.getActive();
  const clientsSheet = getSheetOrThrow_(ss, CLIENTS_SHEET);

  const { appts, clientMap, tomorrow } = getTomorrowAppointments_();
  const todayUnpaid = getTodayUnpaid_();
  
  // Get names of clients already shown in tomorrow's section for dedup
  const tomorrowNames = appts.map(a => nameCase_(a.name));
  const unreliable = getUnreliableAppointments_(tomorrowNames);

  // Only send if there's something to report
  if (appts.length === 0 && todayUnpaid.length === 0 && unreliable.length === 0) return;

  const tomorrowStr = Utilities.formatDate(tomorrow, tz, "EEEE d MMMM");
  let msg = "";

  // ── SECTION 1: Tomorrow's appointments ──
  if (appts.length > 0) {
    msg += `✂️ <b>Tomorrow — ${tomorrowStr}</b>\n`;
    msg += `${appts.length} appointment(s)\n\n`;

    for (const appt of appts) {
      const timeStr = appt.time ? Utilities.formatDate(new Date(appt.time), tz, "HH:mm") : "—";
      const { badge, doNotCut, notes } = getReliabilityInfo_(appt.name, clientMap);
      const priceStr = appt.payment === "Subscription" ? "Sub" : `€${appt.price}`;
      const nameDisplay = doNotCut ? `<u>${appt.name}</u>` : appt.name;
      
      // VIP badge
      const isVip = isVIP_(appt.name, clientsSheet);
      const vipBadge = isVip ? " ⭐" : "";
      
      // Loyalty tier
      const consecutivePaid = getConsecutivePaidCount_(appt.name, clientsSheet);
      const loyaltyLabel = consecutivePaid >= 5 ? " ✅ ELIGIBLE FOR FREE" : "";

      msg += `<b>${timeStr} — ${nameDisplay}${vipBadge}</b>\n`;
      msg += `${priceStr} · ${badge}${loyaltyLabel}\n`;
      
      // Last appointment info
      if (appt.lastAppt) {
        msg += `📋 Last: ${appt.lastAppt.label}`;
        if (appt.lastAppt.notes) {
          msg += ` — "${appt.lastAppt.notes}"`;
        }
        msg += "\n";
      }
      
      if (notes) msg += `📝 <i>${notes}</i>\n`;
      msg += "\n";
    }
  }

  // ── SECTION 2: Today's unpaid ──
  if (todayUnpaid.length > 0) {
    msg += `💸 <b>Unpaid from today (${todayUnpaid.length})</b>\n`;
    for (const appt of todayUnpaid) {
      const timeStr = appt.time ? Utilities.formatDate(new Date(appt.time), tz, "HH:mm") : "—";
      msg += `• ${timeStr} — ${appt.name} (€${appt.price})\n`;
    }
    msg += "\n";
  }

  // ── SECTION 3: Unreliable clients ──
  const twoDayAlerts = unreliable.filter(a => a.daysAhead === 2);
  const oneDayAlerts = unreliable.filter(a => a.daysAhead === 1);

  if (twoDayAlerts.length > 0) {
    const twoDays = addDays_(new Date(), 2);
    const twoDaysStr = Utilities.formatDate(twoDays, tz, "EEEE d MMMM");
    msg += `⚠️ <b>In 2 days — ${twoDaysStr}</b>\n`;
    for (const appt of twoDayAlerts) {
      const timeStr = appt.time ? Utilities.formatDate(new Date(appt.time), tz, "HH:mm") : "—";
      const label = appt.doNotCut ? "⛔ DO NOT CUT" : `${appt.noShows} NS, ${appt.lates} late`;
      msg += `• ${timeStr} — ${appt.name} (${label})\n`;
    }
    msg += "\n";
  }

  if (oneDayAlerts.length > 0) {
    const oneDay = addDays_(new Date(), 1);
    const oneDayStr = Utilities.formatDate(oneDay, tz, "EEEE d MMMM");
    msg += `🚨 <b>TOMORROW — ${oneDayStr}</b>\n`;
    for (const appt of oneDayAlerts) {
      const timeStr = appt.time ? Utilities.formatDate(new Date(appt.time), tz, "HH:mm") : "—";
      const label = appt.doNotCut ? "⛔ DO NOT CUT" : `${appt.noShows} NS, ${appt.lates} late`;
      msg += `• ${timeStr} — ${appt.name} (${label})\n`;
    }
  }

  UrlFetchApp.fetch(`https://api.telegram.org/bot${token}/sendMessage`, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({
      chat_id: chatId,
      text: msg,
      parse_mode: "HTML"
    })
  });
}

function loadClientNotificationMap_(sheet) {
  const map = new Map();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const name = nameCase_(data[i][0]);
    if (!name) continue;
    map.set(name, {
      noShow: Number(data[i][5]) || 0,
      late: Number(data[i][6]) || 0,
      notes: data[i][4] || "",
      doNotCut: data[i][13] === true
    });
  }
  return map;
}

/***************
 * HELPERS
 ***************/
function getStandardServicePrice_(serviceName) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SERVICES_SHEET);
  if (!sheet || !serviceName) return 15;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase() === String(serviceName).toLowerCase()) {
      let p = data[i][1];
      if (typeof p === "string") p = parseFloat(p.replace(/[€\s]/g, "").replace(",", "."));
      return isNaN(p) ? 0 : p;
    }
  }
  return 15;
}

function loadServicePrices_(sheet) {
  const data = sheet.getDataRange().getValues();
  const prices = {};
  for (let i = 1; i < data.length; i++) {
    let p = data[i][1];
    if (typeof p === "string") p = parseFloat(p.replace(/[€\s]/g, "").replace(",", "."));
    if (data[i][0]) prices[String(data[i][0]).toLowerCase()] = isNaN(p) ? 0 : p;
  }
  return prices;
}

function loadClientsIndex_(sheet) {
  const map = new Map();
  let maxId = 0;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const name = nameCase_(data[i][0]);
    const id = data[i][11];
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
  if (sheet.getMaxColumns() < 17) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), 17 - sheet.getMaxColumns());
  }
  sheet.getRange(newRow, 1).setValue(formatted);
  sheet.getRange(newRow, 12).setValue(idx.maxId);
  idx.map.set(formatted, idx.maxId);
  return idx.maxId;
}

function loadActiveSubscriptionsIndex_(sheet) {
  const today = startOfDay_(new Date());
  const data = sheet.getDataRange().getValues();
  const byId = new Map(), byName = new Map();
  for (let i = 1; i < data.length; i++) {
    const name = data[i][0];
    const credits = data[i][4];
    const status = data[i][5];
    const expiryRaw = data[i][6];
    const clientId = data[i][8];
    const expiry = expiryRaw ? new Date(expiryRaw) : null;
    if (status !== "Active" || !expiry || expiry < today) continue;
    if (credits !== "" && Number(credits) <= 0) continue;
    const entry = { start: data[i][3] ? new Date(data[i][3]) : new Date(0) };
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
  const ids = sheet.getRange(1, 12, sheet.getLastRow(), 1).getValues();
  for (let i = 1; i < ids.length; i++) {
    if (ids[i][0]) map.set(String(ids[i][0]), i + 1);
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

function safeAlert_(m) { try { SpreadsheetApp.getUi().alert(m); } catch (e) {} }

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

function testTelegramAPI() {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty("TELEGRAM_BOT_TOKEN");
  const chatId = props.getProperty("TELEGRAM_CHAT_ID");

  Logger.log("Token: " + token);
  Logger.log("Chat ID: " + chatId);

  if (!token || !chatId) {
    Logger.log("ERROR: Token or Chat ID is missing!");
    return;
  }

  const url = `https://api.telegram.org/bot${token}/sendMessage`;
  Logger.log("Calling URL: " + url);

  try {
    const response = UrlFetchApp.fetch(url, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({
        chat_id: chatId,
        text: "🔧 Test from Apps Script"
      })
    });
    Logger.log("Response: " + response.getContentText());
  } catch (e) {
    Logger.log("ERROR: " + e);
  }
}

function debugTomorrow() {
  const result = getTomorrowAppointments_();
  Logger.log("Number of appointments tomorrow: " + result.appts.length);
  Logger.log("Appointments: " + JSON.stringify(result.appts));
  Logger.log("Tomorrow's date: " + result.tomorrow);
}

/**
 * INCREMENTAL SYNC
 * SETUP: Apps Script editor → Services (+) → Google Calendar API → Add
 * Then run: ✂️ Barber Tools → Setup Incremental Sync (5 min)
 * 
 * How it works:
 *   First run  → full sync, stores a "sync token"
 *   Every 5min → sends token, Google returns ONLY what changed
 *   Nothing changed → exits in ~0.4 sec (barely uses quota)
 *   Token expires weekly → auto full sync + new token
 */
function setupIncrementalSync() {
  removeIncrementalSync();
  PropertiesService.getScriptProperties().deleteProperty("CALENDAR_SYNC_TOKEN");
  ScriptApp.newTrigger("syncCalendarIncremental_")
    .timeBased().everyMinutes(5).create();
  syncCalendarIncremental_(); // first run to get initial token
  safeAlert_("✅ Incremental sync running every 5 minutes.\n\nIf you see a Calendar error, go to:\nServices (+) in the left panel → Google Calendar API → Add");
}

function removeIncrementalSync() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "syncCalendarIncremental_") {
      ScriptApp.deleteTrigger(t);
    }
  });
  PropertiesService.getScriptProperties().deleteProperty("CALENDAR_SYNC_TOKEN");
  safeAlert_("✅ Incremental sync stopped.");
}

function syncCalendarIncremental_() {
  // Prevent concurrent runs (web app + trigger firing at the same time)
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(0)) {
    Logger.log("Sync already running, skipping.");
    return;
  }
  try {
    const props = PropertiesService.getScriptProperties();
    const cal   = getCalendarOrThrow_();
    const calId = cal.getId();
    const tz    = cal.getTimeZone();

    let syncToken = props.getProperty("CALENDAR_SYNC_TOKEN");
    let items  = [];
    let newToken = null;

    // ── Try incremental (fast path) ──────────────────────────────────
    if (syncToken) {
      try {
        let resp = Calendar.Events.list(calId, { syncToken, singleEvents: true, fields: 'items(id,iCalUID,status,summary,description,start),nextSyncToken,nextPageToken' });
        items = resp.items || [];
        while (resp.nextPageToken) {
          resp = Calendar.Events.list(calId, { pageToken: resp.nextPageToken });
          items = items.concat(resp.items || []);
        }
        newToken = resp.nextSyncToken;
      } catch (e) {
        // 410 = token expired → fall through to full sync
        syncToken = null;
        props.deleteProperty("CALENDAR_SYNC_TOKEN");
      }
    }

    // ── Full sync (first run or token expired) ───────────────────────
    if (!syncToken) {
      const now  = new Date();
      let resp = Calendar.Events.list(calId, {
        timeMin: startOfDay_(addDays_(now, -DAYS_BACK)).toISOString(),
        timeMax: endOfDay_(addDays_(now,  DAYS_FORWARD)).toISOString(),
        singleEvents: true,
        maxResults: 2500
      });
      items = resp.items || [];
      while (resp.nextPageToken) {
        resp = Calendar.Events.list(calId, { pageToken: resp.nextPageToken });
        items = items.concat(resp.items || []);
      }
      newToken = resp.nextSyncToken;
    }

    // ── Nothing changed — save token and exit ────────────────────────
    if (items.length === 0) {
      if (newToken) props.setProperty("CALENDAR_SYNC_TOKEN", newToken);
      return;
    }

    // ── Something changed — process only those events ────────────────
    const ss          = SpreadsheetApp.getActive();
    const apptSheet   = getSheetOrThrow_(ss, APPOINTMENTS_SHEET);
    const clientsSheet = getSheetOrThrow_(ss, CLIENTS_SHEET);
    const servicesSheet = getSheetOrThrow_(ss, SERVICES_SHEET);
    const subsSheet   = getSheetOrThrow_(ss, SUBSCRIPTIONS_SHEET);
    const today       = startOfDay_(new Date());

    const ctx = {
      ss, cal,
      appointmentsSheet: apptSheet,
      clientsSheet, servicesSheet,
      subscriptionsSheet: subsSheet,
      calTz: tz,
      servicePrices:  loadServicePrices_(servicesSheet),
      clientsIndex:   loadClientsIndex_(clientsSheet),
      subsIndex:      loadActiveSubscriptionsIndex_(subsSheet),
      apptIndex:      loadAppointmentEventIdIndex_(apptSheet),
      startDate:      startOfDay_(addDays_(new Date(), -DAYS_BACK)),
      endDate:        endOfDay_(addDays_(new Date(),  DAYS_FORWARD))
    };

    const newRowsABC = [], newRowsDToM = [];
    let hasChanges = false;
    const newEventIds = new Set();

    for (const item of items) {
      const eventId = item.iCalUID ? String(item.iCalUID) : String(item.id);

      // Deleted event
      if (item.status === "cancelled") {
        const rowIdx = ctx.apptIndex.get(eventId);
        if (!rowIdx || rowIdx === -1) continue;
        const statusCell = apptSheet.getRange(rowIdx, 6);
        const s = statusCell.getValue();
        if (s !== "Paid" && s !== "No Show" && s !== "Cancelled") {
          statusCell.setValue("Cancelled");
          apptSheet.getRange(rowIdx, 8).setValue(false);
          hasChanges = true;
        }
        continue;
      }

      const parsed = parseEventTitle_(item.summary || "");
      if (!parsed) continue;

      const clientId  = getOrCreateClientId_(clientsSheet, ctx.clientsIndex, parsed.clientName);
      const hasCredits = hasActiveCredits_(ctx.subsIndex, parsed.clientName, clientId);
      const start     = new Date(item.start.dateTime || item.start.date);
      const isAllDay  = !item.start.dateTime;
      const { dateCell, timeCell, ymd, hm } = toSheetDateTime_(start, tz, isAllDay);

      const serviceLower = parsed.service.toLowerCase();
      const isSubSale    = serviceLower === "monthly subscription";
      let price, payment, serviceToWrite = parsed.service;

      if (isSubSale) {
        serviceToWrite = "Haircut"; price = 0; payment = "Subscription";
      } else if (hasCredits) {
        price = 0; payment = "Subscription";
      } else {
        price = ctx.servicePrices[serviceLower] ?? 0; payment = "";
      }

      const isFuture     = new Date(dateCell) >= today;
      const initialStatus = payment === "Subscription" ? "Paid" : (isFuture ? "Upcoming" : "Not Paid");
      const existingRow  = ctx.apptIndex.get(eventId);

      if (!existingRow) {
        // New appointment
        newRowsABC.push([dateCell, timeCell, ""]);
        newRowsDToM.push([price, payment, initialStatus, "", false,
          item.description || "", serviceToWrite, clientId, eventId, parsed.clientName]);
        if (isSubSale) createSubscriptionEntry_(ctx, parsed.clientName, clientId, dateCell);
        ctx.apptIndex.set(eventId, -1);
        hasChanges = true;

      } else if (existingRow !== -1) {
        // Updated appointment — refresh time / notes / service
        const rowVals   = apptSheet.getRange(existingRow, 1, 1, 13).getValues()[0];
        const oldDate   = new Date(rowVals[0]);
        const oldTime   = rowVals[1] ? hm_(new Date(rowVals[1]), tz) : "";
        const changed   = ymd !== ymd_(oldDate, tz) ||
                          hm  !== oldTime ||
                          (item.description || "") !== String(rowVals[8] || "") ||
                          serviceToWrite !== String(rowVals[9] || "");
        if (changed) {
          const formula = `=XLOOKUP(K${existingRow}; Clients!L:L; Clients!A:A; M${existingRow})`;
          apptSheet.getRange(existingRow, 1, 1, 3).setValues([[dateCell, timeCell, formula]]);
          apptSheet.getRange(existingRow, 9, 1, 5).setValues([[
            item.description || "", serviceToWrite, clientId, eventId, parsed.clientName
          ]]);
        }
      }
    }

    // Write new rows
    if (newRowsABC.length > 0) {
      const startRow = apptSheet.getLastRow() + 1;
      for (let i = 0; i < newRowsABC.length; i++) {
        const r = startRow + i;
        newRowsABC[i][2] = `=XLOOKUP(K${r}; Clients!L:L; Clients!A:A; M${r})`;
      }
      apptSheet.getRange(startRow, 1, newRowsABC.length, 3).setValues(newRowsABC);
      apptSheet.getRange(startRow, 4, newRowsDToM.length, 10).setValues(newRowsDToM);
    }

    updateUpcomingToNotPaid_(apptSheet);
    updateConsecutivePaidCounts_(ctx);
    sortAndHideAppointments_(apptSheet);

    if (newToken) props.setProperty("CALENDAR_SYNC_TOKEN", newToken);
    if (hasChanges) sendSyncNotification_(ctx, newEventIds);

  } catch (e) {
    Logger.log("Incremental sync error: " + e.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * CLEANUP DUPLICATES
 * Run this once to remove the duplicate rows created by the ID mismatch bug.
 * Safe: only deletes rows where Payment is empty AND status is Not Paid/Upcoming
 * AND another row exists for the same client at the same date+time.
 */
function cleanupDuplicates() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(APPOINTMENTS_SHEET);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { safeAlert_("No data found."); return; }

  const data = sheet.getRange(2, 1, lastRow - 1, 13).getValues();

  // Group rows by date + time + name
  const groups = new Map();
  for (let i = 0; i < data.length; i++) {
    const row     = data[i];
    const dateStr = row[0] ? new Date(row[0]).toDateString() : "";
    const timeStr = row[1] ? String(row[1])                  : "";
    const name    = String(row[2] || row[12] || "").trim().toLowerCase();
    const payment = String(row[4] || "").trim();
    const status  = String(row[5] || "").trim();

    if (!dateStr || !name) continue;

    const key = `${dateStr}|${timeStr}|${name}`;
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push({ sheetRow: i + 2, payment, status });
  }

  // Find rows to delete: duplicates with no payment + Not Paid or Upcoming status
  const toDelete = [];
  for (const rows of groups.values()) {
    if (rows.length <= 1) continue;

    // Keep any row that has a payment or a meaningful status
    const hasGoodRow = rows.some(r =>
      r.payment !== "" ||
      (r.status !== "Not Paid" && r.status !== "Upcoming")
    );
    if (!hasGoodRow) continue;

    for (const r of rows) {
      if (r.payment === "" && (r.status === "Not Paid" || r.status === "Upcoming")) {
        toDelete.push(r.sheetRow);
      }
    }
  }

  if (toDelete.length === 0) {
    safeAlert_("✅ No duplicates found — sheet looks clean!");
    return;
  }

  // Delete from bottom up so row indices stay valid
  toDelete.sort((a, b) => b - a);
  for (const rowIdx of toDelete) {
    sheet.deleteRow(rowIdx);
  }

  safeAlert_(`✅ Removed ${toDelete.length} duplicate rows.\n\nNow run: ✂️ Barber Tools → Setup Incremental Sync`);
}

function doGet(e) {
  try {
    syncCalendarIncremental_(); // shares token with 5-min trigger — no double-processing
    return ContentService
      .createTextOutput("✅ Sync complete!")
      .setMimeType(ContentService.MimeType.TEXT);
  } catch (err) {
    return ContentService
      .createTextOutput("❌ Error: " + err.message)
      .setMimeType(ContentService.MimeType.TEXT);
  }
}
