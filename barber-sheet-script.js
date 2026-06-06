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
const NOTIFICATION_MODE = "telegram";

// Layout constants — table header is in HEADER_ROW, data starts at DATA_ROW.
// Row 1 and column A are spacers (empty). Tables start at B2.
const HEADER_ROW = 2;
const DATA_ROW   = 3;

// Expected header names in HEADER_ROW (must match cell text exactly, case-sensitive):
//
// Appointments:  Date | Time | Name | Price | Payment | Status | Tips | Late | Notes | Service | ClientID | EventID | Cached Name
// Clients:       Name | Favourite Service | Last Visit | Social Media | Notes | No Show | Late | Referral | Total Visits | Total Tips | Total Spent | ClientID | First Visit | Do Not Cut | Consecutive Paid | VIP
// Services:      Service | Price
// Subscriptions: Name | Price | Type | Start Date | Credits | Status | Expiry | Notes | ClientID

/***************
 * COLUMN DISCOVERY
 * getSheetCols_ reads HEADER_ROW and returns {headerName: 1-based col number}.
 * colLetter_ converts a 1-based column number to an A1 letter (e.g. 13 → "M").
 * All data access uses these maps instead of hardcoded numbers.
 ***************/

function getSheetCols_(sheet) {
  const headers = sheet.getRange(HEADER_ROW, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  for (let c = 0; c < headers.length; c++) {
    const name = String(headers[c]).trim();
    if (name) map[name] = c + 1; // 1-based column number
  }
  return map;
}

function colLetter_(n) {
  let s = "";
  while (n > 0) {
    s = String.fromCharCode(65 + (n - 1) % 26) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

// Builds the XLOOKUP formula that resolves a client name from the Appointments sheet.
// Uses dynamic column letters so it stays correct if columns move.
function buildNameFormula_(ctx, row) {
  const clientIdLet      = colLetter_(ctx.apptCols["ClientID"]);
  const cachedNameLet    = colLetter_(ctx.apptCols["Cached Name"]);
  const clientsIdLet     = colLetter_(ctx.clientCols["ClientID"]);
  const clientsNameLet   = colLetter_(ctx.clientCols["Name"]);
  return `=XLOOKUP(${clientIdLet}${row}; Clients!${clientsIdLet}:${clientsIdLet}; Clients!${clientsNameLet}:${clientsNameLet}; ${cachedNameLet}${row})`;
}

/***************
 * MENU & TRIGGERS
 ***************/

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("✂️ Barber Tools")
    .addItem("🔄 Sync Now", "syncCalendarIncremental_")
    .addToUi();
}

// INSTALLABLE TRIGGER — linked to 'On edit' in Script Settings
function processSheetChanges(e) {
  const range = e && e.range;
  if (!range) return;
  const sheet = range.getSheet();
  const sheetName = sheet.getName();

  // 1. Dashboard mobile sync button (C3)
  if (sheetName === "Dashboard" && range.getRow() === 3 && range.getColumn() === 3) {
    if (range.getValue() === true) {
      range.setValue(false);
      SpreadsheetApp.flush();
      syncCalendarIncremental_();
    }
    return;
  }

  // 2. Appointments sheet logic
  if (sheetName === APPOINTMENTS_SHEET && range.getRow() > HEADER_ROW) {
    const aC = getSheetCols_(sheet); // one header read covers all column lookups below
    const row = range.getRow();
    const col = range.getColumn();

    // A. Payment changed
    if (col === aC["Payment"]) {
      const paymentVal  = range.getValue();
      const statusRange = sheet.getRange(row, aC["Status"]);
      const currentStatus = statusRange.getValue();

      if (currentStatus !== "No Show" && currentStatus !== "Cancelled") {
        if (paymentVal === "Cash" || paymentVal === "Tikkie" || paymentVal === "Subscription" || paymentVal === "Free") {
          statusRange.setValue("Paid");
        } else if (paymentVal === "" && currentStatus === "Paid") {
          const dateVal = sheet.getRange(row, aC["Date"]).getValue();
          const isUpcoming = dateVal && new Date(dateVal) >= startOfDay_(new Date());
          statusRange.setValue(isUpcoming ? "Upcoming" : "Not Paid");
        }
      }

      const priceRange   = sheet.getRange(row, aC["Price"]);
      const currentPrice = priceRange.getValue();
      const serviceName  = String(sheet.getRange(row, aC["Service"]).getValue()).toLowerCase();

      if (paymentVal === "Subscription" || paymentVal === "Free") {
        priceRange.setValue(0);
      } else if (paymentVal === "" && currentPrice === 0) {
        priceRange.setValue(getStandardServicePrice_(serviceName));
      }
    }

    // B. Status changed
    if (col === aC["Status"]) {
      const statusVal = range.getValue();
      if (statusVal === "No Show" || statusVal === "Cancelled") {
        sheet.getRange(row, aC["Late"]).setValue(false);
      }
    }

    // C. Recalculate client stats when any stat-affecting column changes
    const statCols = new Set([aC["Price"], aC["Payment"], aC["Status"], aC["Tips"], aC["Late"], aC["ClientID"]]);
    if (statCols.has(col)) {
      const clientsSheet = sheet.getParent().getSheetByName(CLIENTS_SHEET);
      const minCtx = {
        appointmentsSheet: sheet,
        clientsSheet,
        apptCols:   aC,
        clientCols: getSheetCols_(clientsSheet)
      };
      updateClientStats_(minCtx);
      updateNoShowLateCounts_(minCtx);
      updateConsecutivePaidCounts_(minCtx);
    }
  }

  // 3. Client name auto-formatting
  if (sheetName === CLIENTS_SHEET && range.getRow() > HEADER_ROW) {
    const cC = getSheetCols_(sheet);
    if (range.getColumn() === cC["Name"]) {
      const v = range.getValue();
      const fixed = nameCase_(v);
      if (fixed !== v) range.setValue(fixed);
    }
  }
}

/***************
 * MAIN SYNC ENGINES
 ***************/

function syncCalendarToSheets(showNotification = true) {
  const ctx = prepareContext_();
  const now = new Date();
  const startDate = startOfDay_(addDays_(now, -DAYS_BACK));
  const endDate   = endOfDay_(addDays_(now, DAYS_FORWARD));
  const events = ctx.calendar.getEvents(startDate, endDate);

  const counts = upsertEvents_(events, { ...ctx, startDate, endDate });
  updateUpcomingToNotPaid_(ctx);
  const apptLastRow = ctx.appointmentsSheet.getLastRow();
  const apptData = loadApptData_(ctx.appointmentsSheet, apptLastRow);
  updateConsecutivePaidCounts_(ctx, apptData);
  updateNoShowLateCounts_(ctx, apptData);
  updateClientStats_(ctx, apptData);
  sortAndHideAppointments_(ctx);

  if (showNotification) {
    safeAlert_(`Sync Complete!\n\n+ ${counts.newCount} New\n~ ${counts.updatedCount} Updated\n- ${counts.cancelledCount} Cancelled`);
    if (counts.newCount > 0 || counts.cancelledCount > 0) {
      sendSyncNotification_(ctx, counts.newEventIds);
    }
  }
}

function syncThisYear(showNotification = true) {
  const ctx = prepareContext_();
  const now  = new Date();
  const year = now.getFullYear();
  const startDate = new Date(year, 0, 1);
  const endDate   = new Date(year, 11, 31, 23, 59, 59, 999);

  const events = ctx.calendar.getEvents(startDate, endDate);
  const counts = upsertEvents_(events, { ...ctx, startDate, endDate });
  updateUpcomingToNotPaid_(ctx);
  const apptLastRow = ctx.appointmentsSheet.getLastRow();
  const apptData = loadApptData_(ctx.appointmentsSheet, apptLastRow);
  updateConsecutivePaidCounts_(ctx, apptData);
  updateNoShowLateCounts_(ctx, apptData);
  updateClientStats_(ctx, apptData);
  sortAndHideAppointments_(ctx);

  if (showNotification) {
    safeAlert_(`Sync Complete!\n\n+ ${counts.newCount} New\n~ ${counts.updatedCount} Updated\n- ${counts.cancelledCount} Cancelled`);
    if (counts.newCount > 0 || counts.cancelledCount > 0) {
      sendSyncNotification_(ctx, counts.newEventIds);
    }
  }
}

// Reads all appointment rows into a 2D array starting from col 1.
// Array index for any field = ctx.apptCols["FieldName"] - 1.
function loadApptData_(sheet, lastRow) {
  if (lastRow < DATA_ROW) return [];
  return sheet.getRange(DATA_ROW, 1, lastRow - DATA_ROW + 1, sheet.getLastColumn()).getValues();
}

function prepareContext_() {
  const ss  = SpreadsheetApp.getActive();
  const cal = getCalendarOrThrow_();

  const apptSheet    = getSheetOrThrow_(ss, APPOINTMENTS_SHEET);
  const clientsSheet = getSheetOrThrow_(ss, CLIENTS_SHEET);
  const servicesSheet = getSheetOrThrow_(ss, SERVICES_SHEET);
  const subsSheet    = getSheetOrThrow_(ss, SUBSCRIPTIONS_SHEET);

  // Read column positions from header row — the single source of truth for layout
  const apptCols   = getSheetCols_(apptSheet);
  const clientCols = getSheetCols_(clientsSheet);
  const svcCols    = getSheetCols_(servicesSheet);
  const subsCols   = getSheetCols_(subsSheet);

  return {
    ss,
    appointmentsSheet: apptSheet,
    clientsSheet,
    servicesSheet,
    subscriptionsSheet: subsSheet,
    calendar: cal,
    calTz:    cal.getTimeZone(),
    apptCols, clientCols, svcCols, subsCols,
    servicePrices: loadServicePrices_(servicesSheet, svcCols),
    clientsIndex:  loadClientsIndex_(clientsSheet, clientCols),
    subsIndex:     loadActiveSubscriptionsIndex_(subsSheet, subsCols),
    apptIndex:     loadAppointmentEventIdIndex_(apptSheet, apptCols)
  };
}

function upsertEvents_(events, ctx) {
  const today = startOfDay_(new Date());
  const ac    = ctx.apptCols;
  const ncols = ctx.appointmentsSheet.getLastColumn();
  let newCount = 0, updatedCount = 0, cancelledCount = 0;
  const validEventIds = new Set();
  const newEventIds   = new Set();
  const newRows       = []; // full-width row arrays for new appointments

  for (const event of events) {
    const eventId = String(event.getId());
    validEventIds.add(eventId);
    const parsed = parseEventTitle_(event.getTitle());
    if (!parsed) continue;

    const existing   = ctx.apptIndex.get(eventId);
    const clientId   = getOrCreateClientId_(ctx.clientsSheet, ctx.clientsIndex, parsed.clientName);
    const hasCredits = hasActiveCredits_(ctx.subsIndex, parsed.clientName, clientId);
    const start      = event.getStartTime();
    const { dateCell, timeCell, ymd, hm } = toSheetDateTime_(start, ctx.calTz, event.isAllDayEvent());

    const serviceLower      = parsed.service.toLowerCase();
    const isSubscriptionSale = serviceLower === "monthly subscription";
    let price, payment, serviceToWrite = parsed.service;

    if (isSubscriptionSale) {
      serviceToWrite = "Haircut"; price = 0; payment = "Subscription";
    } else if (hasCredits) {
      price = 0; payment = "Subscription";
    } else {
      price = ctx.servicePrices[serviceLower] ?? 0; payment = "";
    }

    const isFuture     = new Date(dateCell) >= today;
    const initialStatus = (payment === "Subscription") ? "Paid" : (isFuture ? "Upcoming" : "Not Paid");

    if (!existing) {
      // --- NEW APPOINTMENT ---
      const rowArr = new Array(ncols).fill("");
      rowArr[ac["Date"] - 1]       = dateCell;
      rowArr[ac["Time"] - 1]       = timeCell;
      // Name formula: filled in below once we know the target row number
      rowArr[ac["Price"] - 1]      = price;
      rowArr[ac["Payment"] - 1]    = payment;
      rowArr[ac["Status"] - 1]     = initialStatus;
      rowArr[ac["Tips"] - 1]       = "";
      rowArr[ac["Late"] - 1]       = false;
      rowArr[ac["Notes"] - 1]      = event.getDescription() || "";
      rowArr[ac["Service"] - 1]    = serviceToWrite;
      rowArr[ac["ClientID"] - 1]   = clientId;
      rowArr[ac["EventID"] - 1]    = eventId;
      rowArr[ac["Cached Name"] - 1] = parsed.clientName;
      newRows.push(rowArr);

      if (isSubscriptionSale) createSubscriptionEntry_(ctx, parsed.clientName, clientId, dateCell);
      ctx.apptIndex.set(eventId, -1);
      newEventIds.add(eventId);
      newCount++;

    } else {
      // --- EXISTING APPOINTMENT ---
      if (existing === -1) continue;

      const existingRow = existing.row;
      const rowVals     = existing.data; // full-row array read from col 1

      const oldDateVal = new Date(rowVals[ac["Date"] - 1]);
      const oldStatus  = rowVals[ac["Status"] - 1];
      const oldPayment = String(rowVals[ac["Payment"] - 1] || "");
      const oldTimeRaw = rowVals[ac["Time"] - 1];
      const oldTimeStr = oldTimeRaw ? hm_(new Date(oldTimeRaw), ctx.calTz) : "";
      const oldService = String(rowVals[ac["Service"] - 1] || "");
      const oldName    = String(rowVals[ac["Name"] - 1] || "");

      // Flip Upcoming → Not Paid if date has passed
      if (oldStatus === "Upcoming" && oldDateVal < today) {
        ctx.appointmentsSheet.getRange(existingRow, ac["Status"]).setValue("Not Paid");
      }

      // Seamless subscription conversion
      if (isSubscriptionSale && oldPayment !== "Subscription") {
        ctx.appointmentsSheet.getRange(existingRow, ac["Price"]).setValue(0);
        ctx.appointmentsSheet.getRange(existingRow, ac["Payment"]).setValue("Subscription");
        ctx.appointmentsSheet.getRange(existingRow, ac["Status"]).setValue("Paid");
        createSubscriptionEntry_(ctx, parsed.clientName, clientId, dateCell);
      }

      const changed =
        ymd !== ymd_(oldDateVal, ctx.calTz) ||
        hm  !== oldTimeStr ||
        (event.getDescription() || "") !== String(rowVals[ac["Notes"] - 1] || "") ||
        serviceToWrite !== oldService ||
        parsed.clientName !== oldName;

      if (changed) {
        let rowPrice = rowVals[ac["Price"] - 1];
        const isUnpaid = oldStatus === "Not Paid" || oldStatus === "Upcoming";
        if (isUnpaid && new Date(dateCell) >= today) {
          rowPrice = (isSubscriptionSale || hasCredits) ? 0 : (ctx.servicePrices[serviceLower] ?? rowPrice);
        }

        // Write the full row at once: copy existing data, overwrite changed fields
        const updatedRow = rowVals.slice();
        updatedRow[ac["Date"] - 1]       = dateCell;
        updatedRow[ac["Time"] - 1]       = timeCell;
        updatedRow[ac["Name"] - 1]       = buildNameFormula_(ctx, existingRow);
        updatedRow[ac["Price"] - 1]      = isSubscriptionSale ? 0 : rowPrice;
        updatedRow[ac["Notes"] - 1]      = event.getDescription() || "";
        updatedRow[ac["Service"] - 1]    = serviceToWrite;
        updatedRow[ac["ClientID"] - 1]   = clientId;
        updatedRow[ac["EventID"] - 1]    = eventId;
        updatedRow[ac["Cached Name"] - 1] = parsed.clientName;
        ctx.appointmentsSheet.getRange(existingRow, 1, 1, updatedRow.length).setValues([updatedRow]);
        updatedCount++;
      }
    }
  }

  // Mark cancelled appointments (calendar events deleted)
  ctx.apptIndex.forEach((entry, eId) => {
    if (entry === -1) return;
    if (validEventIds.has(eId)) return;

    const rowIdx  = entry.row;
    const dateVal = entry.data[ac["Date"] - 1];
    if (!dateVal) return;
    const rowDate = new Date(dateVal);
    if (rowDate < ctx.startDate || rowDate > ctx.endDate) return;

    const currentStatus = String(entry.data[ac["Status"] - 1] || "");
    if (currentStatus === "Paid" || currentStatus === "No Show" || currentStatus === "Cancelled") return;

    ctx.appointmentsSheet.getRange(rowIdx, ac["Status"]).setValue("Cancelled");
    ctx.appointmentsSheet.getRange(rowIdx, ac["Late"]).setValue(false);
    cancelledCount++;
  });

  // Write all new rows in one batch
  if (newRows.length > 0) {
    const startRow = ctx.appointmentsSheet.getLastRow() + 1;
    for (let i = 0; i < newRows.length; i++) {
      newRows[i][ac["Name"] - 1] = buildNameFormula_(ctx, startRow + i);
    }
    ctx.appointmentsSheet.getRange(startRow, 1, newRows.length, ncols).setValues(newRows);
  }

  return { newCount, updatedCount, cancelledCount, newEventIds };
}

// Flip any "Upcoming" appointments whose date has passed to "Not Paid"
function updateUpcomingToNotPaid_(ctx) {
  const sheet   = ctx.appointmentsSheet;
  const ac      = ctx.apptCols;
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_ROW) return;
  const today = startOfDay_(new Date());
  const data  = sheet.getRange(DATA_ROW, 1, lastRow - DATA_ROW + 1, sheet.getLastColumn()).getValues();
  for (let i = 0; i < data.length; i++) {
    const dateVal = data[i][ac["Date"] - 1];
    const status  = data[i][ac["Status"] - 1];
    if (!dateVal || status !== "Upcoming") continue;
    if (new Date(dateVal) < today) {
      sheet.getRange(i + DATA_ROW, ac["Status"]).setValue("Not Paid");
    }
  }
}

// Update ConsecutivePaid column in Clients sheet
function updateConsecutivePaidCounts_(ctx, apptData) {
  const clientsSheet = ctx.clientsSheet;
  const apptSheet    = ctx.appointmentsSheet;
  const ac = ctx.apptCols;
  const cc = ctx.clientCols;
  const lastRow = clientsSheet.getLastRow();
  if (lastRow < DATA_ROW) return;

  const clientData = clientsSheet.getRange(DATA_ROW, 1, lastRow - DATA_ROW + 1, clientsSheet.getLastColumn()).getValues();
  if (!apptData) apptData = loadApptData_(apptSheet, apptSheet.getLastRow());

  const n     = clientData.length;
  const oVals = new Array(n);

  for (let i = 0; i < n; i++) {
    const clientId = clientData[i][cc["ClientID"] - 1];
    if (!clientId) { oVals[i] = [0]; continue; }

    const clientAppts = [];
    for (const row of apptData) {
      if (String(row[ac["ClientID"] - 1]) === String(clientId)) {
        clientAppts.push({
          date:    row[ac["Date"] - 1],
          status:  row[ac["Status"] - 1],
          payment: row[ac["Payment"] - 1],
          late:    row[ac["Late"] - 1]
        });
      }
    }
    clientAppts.sort((a, b) => new Date(b.date) - new Date(a.date));

    let consecutivePaid = 0;
    for (const appt of clientAppts) {
      if (appt.status === "No Show" || appt.late === true) break;
      const isPaid = appt.payment === "Cash" || appt.payment === "Tikkie" ||
                     appt.payment === "Subscription" || appt.payment === "Free" ||
                     String(appt.status).startsWith("Free");
      if (isPaid) consecutivePaid++;
      else break;
    }
    oVals[i] = [consecutivePaid];
  }

  clientsSheet.getRange(DATA_ROW, cc["Consecutive Paid"], n, 1).setValues(oVals);
}

// Update TotalVisits, TotalTips, TotalSpent, LastVisit, FirstVisit in Clients sheet
function updateClientStats_(ctx, apptData) {
  const clientsSheet = ctx.clientsSheet;
  const apptSheet    = ctx.appointmentsSheet;
  const ac = ctx.apptCols;
  const cc = ctx.clientCols;
  const clientLastRow = clientsSheet.getLastRow();
  if (clientLastRow < DATA_ROW) return;

  const clientData = clientsSheet.getRange(DATA_ROW, 1, clientLastRow - DATA_ROW + 1, clientsSheet.getLastColumn()).getValues();
  if (!apptData) apptData = loadApptData_(apptSheet, apptSheet.getLastRow());

  const lastVisits  = [], totalVisitVals = [], totalTipsVals = [], totalSpentVals = [], firstVisits = [];

  for (let i = 0; i < clientData.length; i++) {
    const clientId = clientData[i][cc["ClientID"] - 1];
    if (!clientId) {
      lastVisits.push([""]); totalVisitVals.push([0]); totalTipsVals.push([0]);
      totalSpentVals.push([0]); firstVisits.push([""]);
      continue;
    }

    let totalVisits = 0, totalTips = 0, totalSpent = 0;
    let firstVisit = null, lastVisit = null;

    for (const row of apptData) {
      if (String(row[ac["ClientID"] - 1]) !== String(clientId)) continue;

      const payment = String(row[ac["Payment"] - 1] || "");
      const status  = String(row[ac["Status"] - 1] || "");
      const isPaidVisit = payment === "Cash" || payment === "Tikkie" ||
                          payment === "Subscription" || payment === "Free" ||
                          status.startsWith("Free");
      if (!isPaidVisit) continue;

      const dateVal = row[ac["Date"] - 1];
      if (dateVal) {
        const d = new Date(dateVal);
        if (!firstVisit || d < firstVisit) firstVisit = d;
        if (!lastVisit  || d > lastVisit)  lastVisit  = d;
      }

      totalVisits++;
      totalTips  += Number(row[ac["Tips"] - 1]) || 0;
      if (payment === "Cash" || payment === "Tikkie" || payment === "Subscription") {
        totalSpent += Number(row[ac["Price"] - 1]) || 0;
      }
    }

    lastVisits.push([lastVisit  || ""]);
    totalVisitVals.push([totalVisits]);
    totalTipsVals.push([totalTips]);
    totalSpentVals.push([totalSpent]);
    firstVisits.push([firstVisit || ""]);
  }

  const n = clientData.length;
  clientsSheet.getRange(DATA_ROW, cc["Last Visit"],      n, 1).setValues(lastVisits);
  clientsSheet.getRange(DATA_ROW, cc["Total Visits"],    n, 1).setValues(totalVisitVals);
  clientsSheet.getRange(DATA_ROW, cc["Total Tips"],      n, 1).setValues(totalTipsVals);
  clientsSheet.getRange(DATA_ROW, cc["Total Spent"],     n, 1).setValues(totalSpentVals);
  clientsSheet.getRange(DATA_ROW, cc["First Visit"],     n, 1).setValues(firstVisits);
}

// Update NoShow and Late count columns in Clients sheet
function updateNoShowLateCounts_(ctx, apptData) {
  const clientsSheet = ctx.clientsSheet;
  const apptSheet    = ctx.appointmentsSheet;
  const ac = ctx.apptCols;
  const cc = ctx.clientCols;
  const clientLastRow = clientsSheet.getLastRow();
  if (clientLastRow < DATA_ROW) return;

  if (!apptData) {
    const apptLastRow = apptSheet.getLastRow();
    if (apptLastRow < DATA_ROW) return;
    apptData = loadApptData_(apptSheet, apptLastRow);
  }

  const cutoff = new Date();
  cutoff.setFullYear(cutoff.getFullYear() - 1);

  const clientData  = clientsSheet.getRange(DATA_ROW, 1, clientLastRow - DATA_ROW + 1, clientsSheet.getLastColumn()).getValues();
  const n           = clientData.length;
  const noShowVals  = new Array(n);
  const lateVals    = new Array(n);

  for (let i = 0; i < n; i++) {
    const clientId = clientData[i][cc["ClientID"] - 1];
    if (!clientId) { noShowVals[i] = [0]; lateVals[i] = [0]; continue; }

    let noShows = 0, lates = 0;
    for (const row of apptData) {
      if (String(row[ac["ClientID"] - 1]) !== String(clientId)) continue;
      const dateVal = row[ac["Date"] - 1];
      if (!dateVal || new Date(dateVal) < cutoff) continue;
      if (row[ac["Status"] - 1] === "No Show") noShows++;
      if (row[ac["Late"] - 1] === true) lates++;
    }
    noShowVals[i] = [noShows];
    lateVals[i]   = [lates];
  }

  clientsSheet.getRange(DATA_ROW, cc["No Show"], n, 1).setValues(noShowVals);
  clientsSheet.getRange(DATA_ROW, cc["Late"],   n, 1).setValues(lateVals);
}

function sortAndHideAppointments_(ctx) {
  const sheet   = ctx.appointmentsSheet;
  const ac      = ctx.apptCols;
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_ROW) return;

  const numCols    = sheet.getLastColumn();
  const numRows    = lastRow - DATA_ROW + 1;
  // Sort column numbers are relative to the range's first column (col 1)
  const dateRelCol = ac["Date"]; // col index from col 1: colNum - 1 + 1 = colNum
  const timeRelCol = ac["Time"];
  sheet.getRange(DATA_ROW, 1, numRows, numCols).sort([
    { column: dateRelCol, ascending: false },
    { column: timeRelCol, ascending: false }
  ]);

  // Read just the Date column to decide which rows to hide
  const dates  = sheet.getRange(DATA_ROW, ac["Date"], numRows, 1).getValues();
  const today  = new Date();
  const cutoff = new Date(today);
  cutoff.setDate(today.getDate() - HIDE_OLDER_THAN_DAYS);

  let hideStartRow = -1;
  for (let i = 0; i < dates.length; i++) {
    const raw = dates[i][0];
    if (!raw) continue;
    const d = new Date(raw);
    if (isNaN(d.getTime())) continue;
    if (d < cutoff) { hideStartRow = i + DATA_ROW; break; }
  }

  sheet.showRows(DATA_ROW, numRows);
  if (hideStartRow !== -1) {
    const rowsToHide = (lastRow + 1) - hideStartRow;
    if (rowsToHide > 0) sheet.hideRows(hideStartRow, rowsToHide);
  }
}

function createSubscriptionEntry_(ctx, clientName, clientId, startDate) {
  const subsSheet    = ctx.subscriptionsSheet;
  const ac  = ctx.apptCols;
  const sc  = ctx.subsCols;
  const cc  = ctx.clientCols;
  const monthlyPrice = ctx.servicePrices["monthly subscription"] ?? 40;
  const name = nameCase_(clientName);

  if (ctx.subsIndex.byName.has(name) || (clientId && ctx.subsIndex.byId.has(String(clientId)))) return;

  const lastRow = subsSheet.getLastRow();
  if (lastRow >= DATA_ROW) {
    const checkStart = Math.max(DATA_ROW, lastRow - 20);
    const checkCount = lastRow - checkStart + 1;
    const checkData  = subsSheet.getRange(checkStart, 1, checkCount, subsSheet.getLastColumn()).getValues();
    const startYMD   = ymd_(startDate instanceof Date ? startDate : new Date(startDate), ctx.calTz);
    for (const r of checkData) {
      const rowName      = nameCase_(String(r[sc["Name"] - 1] || ""));
      const rowStartDate = r[sc["Start Date"] - 1];
      if (rowName === name && rowStartDate && ymd_(new Date(rowStartDate), ctx.calTz) === startYMD) return;
    }
  }

  const r = lastRow + 1;

  // XLOOKUP: resolve client name from Clients sheet using ClientID
  const clientIdLetSubs    = colLetter_(sc["ClientID"]);
  const clientIdLetClients = colLetter_(cc["ClientID"]);
  const nameLetClients     = colLetter_(cc["Name"]);
  const nameFormula = `=XLOOKUP(${clientIdLetSubs}${r}; Clients!${clientIdLetClients}:${clientIdLetClients}; Clients!${nameLetClients}:${nameLetClients}; "")`;

  // Credits formula: counts Subscription-paid appointments since subscription start date
  const clientIdLetAppt  = colLetter_(ac["ClientID"]);
  const paymentLetAppt   = colLetter_(ac["Payment"]);
  const statusLetAppt    = colLetter_(ac["Status"]);
  const dateLetAppt      = colLetter_(ac["Date"]);
  const startDateLetSubs = colLetter_(sc["Start Date"]);
  const expiryLetSubs    = colLetter_(sc["Expiry"]);
  const creditsFormula = `=MAX(0; 4 - COUNTIFS(Appointments!$${clientIdLetAppt}:$${clientIdLetAppt}; ${clientIdLetSubs}${r}; Appointments!$${paymentLetAppt}:$${paymentLetAppt}; "Subscription"; Appointments!$${statusLetAppt}:$${statusLetAppt}; "Paid"; Appointments!$${dateLetAppt}:$${dateLetAppt}; ">="&${startDateLetSubs}${r}; Appointments!$${dateLetAppt}:$${dateLetAppt}; "<="&(${expiryLetSubs}${r} + 21)))`;

  const ncols  = subsSheet.getLastColumn();
  const rowArr = new Array(ncols).fill("");
  rowArr[sc["Name"] - 1]      = nameFormula;
  rowArr[sc["Price"] - 1]     = monthlyPrice;
  rowArr[sc["Start Date"] - 1] = startDate;
  rowArr[sc["Credits"] - 1]   = creditsFormula;
  rowArr[sc["Status"] - 1]    = "Active";
  rowArr[sc["Expiry"] - 1]    = addDays_(startDate, 31);
  rowArr[sc["ClientID"] - 1]  = clientId;

  subsSheet.getRange(r, 1, 1, ncols).setValues([rowArr]);

  // Sort by StartDate descending
  const newLastRow = subsSheet.getLastRow();
  if (newLastRow >= DATA_ROW) {
    subsSheet.getRange(DATA_ROW, 1, newLastRow - DATA_ROW + 1, ncols)
      .sort({ column: sc["Start Date"], ascending: false });
  }

  const entry = { start: startDate };
  ctx.subsIndex.byName.set(name, entry);
  if (clientId) ctx.subsIndex.byId.set(String(clientId), entry);
}

/***************
 * ONE-TIME MIGRATION
 ***************/

function runMigration() {
  const ss    = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(APPOINTMENTS_SHEET);
  const cSheet = ss.getSheetByName(CLIENTS_SHEET);
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_ROW) { Logger.log("No appointment data to migrate."); return; }

  const ac = getSheetCols_(sheet);
  const cc = getSheetCols_(cSheet);

  const data = sheet.getRange(DATA_ROW, 1, lastRow - DATA_ROW + 1, sheet.getLastColumn()).getValues();
  const cachedNames    = [];
  const updatedFormulas = [];

  for (let i = 0; i < data.length; i++) {
    const row           = i + DATA_ROW;
    const currentName   = String(data[i][ac["Name"] - 1] || "");
    const existingCache = String(data[i][ac["Cached Name"] - 1] || "");
    cachedNames.push([existingCache || currentName]);
    updatedFormulas.push([buildNameFormula_({ apptCols: ac, clientCols: cc }, row)]);
  }

  sheet.getRange(DATA_ROW, ac["Cached Name"], cachedNames.length,    1).setValues(cachedNames);
  sheet.getRange(DATA_ROW, ac["Name"],       updatedFormulas.length, 1).setFormulas(updatedFormulas);

  const migMsg = `✅ Migration complete! ${cachedNames.length} rows updated.`;
  Logger.log(migMsg);
  SpreadsheetApp.getActive().toast(migMsg, "Migration", 8);
}

/***************
 * HELPERS
 ***************/

function getStandardServicePrice_(serviceName) {
  const ss    = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SERVICES_SHEET);
  if (!sheet || !serviceName) return 15;
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_ROW) return 15;
  const cols = getSheetCols_(sheet);
  const data = sheet.getRange(DATA_ROW, 1, lastRow - DATA_ROW + 1, sheet.getLastColumn()).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][cols["Service"] - 1]).toLowerCase() === String(serviceName).toLowerCase()) {
      let p = data[i][cols["Price"] - 1];
      if (typeof p === "string") p = parseFloat(p.replace(/[€\s]/g, "").replace(",", "."));
      return isNaN(p) ? 0 : p;
    }
  }
  return 15;
}

function loadServicePrices_(sheet, cols) {
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_ROW) return {};
  const data   = sheet.getRange(DATA_ROW, 1, lastRow - DATA_ROW + 1, sheet.getLastColumn()).getValues();
  const prices = {};
  for (let i = 0; i < data.length; i++) {
    let p = data[i][cols["Price"] - 1];
    if (typeof p === "string") p = parseFloat(p.replace(/[€\s]/g, "").replace(",", "."));
    const svc = data[i][cols["Service"] - 1];
    if (svc) prices[String(svc).toLowerCase()] = isNaN(p) ? 0 : p;
  }
  return prices;
}

// Returns { map: Map<name, id>, maxId: number, cols: colMap }
// cols is stored so getOrCreateClientId_ can write to the right columns without extra API calls.
function loadClientsIndex_(sheet, cols) {
  const map = new Map();
  let maxId = 0;
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_ROW) return { map, maxId, cols };
  const data = sheet.getRange(DATA_ROW, 1, lastRow - DATA_ROW + 1, sheet.getLastColumn()).getValues();
  for (let i = 0; i < data.length; i++) {
    const name = nameCase_(data[i][cols["Name"] - 1]);
    const id   = data[i][cols["ClientID"] - 1];
    if (name) map.set(name, id);
    const n = parseInt(id, 10);
    if (!isNaN(n) && n > maxId) maxId = n;
  }
  return { map, maxId, cols };
}

function getOrCreateClientId_(sheet, idx, name) {
  const formatted = nameCase_(name);
  if (idx.map.has(formatted)) return idx.map.get(formatted);

  idx.maxId += 1;
  const newRow = sheet.getLastRow() + 1;
  sheet.getRange(newRow, idx.cols["Name"]).setValue(formatted);
  sheet.getRange(newRow, idx.cols["ClientID"]).setValue(idx.maxId);
  idx.map.set(formatted, idx.maxId);
  return idx.maxId;
}

function loadActiveSubscriptionsIndex_(sheet, cols) {
  const today   = startOfDay_(new Date());
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_ROW) return { byId: new Map(), byName: new Map() };
  const data  = sheet.getRange(DATA_ROW, 1, lastRow - DATA_ROW + 1, sheet.getLastColumn()).getValues();
  const byId  = new Map(), byName = new Map();
  for (let i = 0; i < data.length; i++) {
    const name      = data[i][cols["Name"] - 1];
    const credits   = data[i][cols["Credits"] - 1];
    const status    = data[i][cols["Status"] - 1];
    const expiryRaw = data[i][cols["Expiry"] - 1];
    const clientId  = data[i][cols["ClientID"] - 1];
    const startRaw  = data[i][cols["Start Date"] - 1];
    const expiry = expiryRaw ? new Date(expiryRaw) : null;
    if (status !== "Active" || !expiry || expiry < today) continue;
    if (credits !== "" && Number(credits) <= 0) continue;
    const entry = { start: startRaw ? new Date(startRaw) : new Date(0) };
    if (clientId) byId.set(String(clientId), entry);
    if (name && !String(name).startsWith("=")) byName.set(nameCase_(name), entry);
  }
  return { byId, byName };
}

function hasActiveCredits_(idx, name, id) {
  return (id && idx.byId.has(String(id))) || (name && idx.byName.has(nameCase_(name)));
}

// Reads appointment event IDs into a map: eventId → { row, data }.
// data is a full-row array (from col 1); use apptCols["Field"] - 1 to index into it.
function loadAppointmentEventIdIndex_(sheet, cols) {
  const map     = new Map();
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_ROW) return map;
  const data = sheet.getRange(DATA_ROW, 1, lastRow - DATA_ROW + 1, sheet.getLastColumn()).getValues();
  for (let i = 0; i < data.length; i++) {
    const eventId = data[i][cols["EventID"] - 1];
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

function safeAlert_(m) { try { SpreadsheetApp.getUi().alert(m); } catch (e) {} }

function nameCase_(v) {
  return String(v ?? "").trim().toLowerCase()
    .split(" ").map(w => w ? w[0].toUpperCase() + w.slice(1) : "").join(" ");
}

function toSheetDateTime_(s, tz, isAllDay) {
  const ymd = Utilities.formatDate(s, tz, "yyyy-MM-dd");
  const hm  = Utilities.formatDate(s, tz, "HH:mm");
  const [y, m, d] = ymd.split("-").map(Number);
  const timeCell = isAllDay ? "" : (() => {
    const [hh, mm] = hm.split(":").map(Number);
    return new Date(1899, 11, 30, hh, mm);
  })();
  return { dateCell: new Date(y, m - 1, d), timeCell, ymd, hm };
}

function startOfDay_(d) { const x = new Date(d); x.setHours(0, 0, 0, 0); return x; }
function endOfDay_(d)   { const x = new Date(d); x.setHours(23, 59, 59, 999); return x; }
function addDays_(d, n) { const x = new Date(d); x.setDate(x.getDate() + n); return x; }
function ymd_(d, tz) { return Utilities.formatDate(d, tz, "yyyy-MM-dd"); }
function hm_(d, tz)  { return Utilities.formatDate(d, tz, "HH:mm"); }

/***************
 * INCREMENTAL SYNC
 * SETUP: Apps Script editor → Services (+) → Google Calendar API → Add
 * Then run setupTriggers() or setupIncrementalSync()
 ***************/

function setupIncrementalSync() {
  removeIncrementalSync();
  PropertiesService.getScriptProperties().deleteProperty("CALENDAR_SYNC_TOKEN");
  ScriptApp.newTrigger("syncCalendarIncremental_").timeBased().everyMinutes(5).create();
  syncCalendarIncremental_();
  const msg = "✅ Incremental sync running every 5 minutes.\n\nIf you see a Calendar error, go to:\nServices (+) in the left panel → Google Calendar API → Add";
  Logger.log(msg);
  SpreadsheetApp.getActive().toast(msg, "Setup", 10);
}

function removeIncrementalSync() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "syncCalendarIncremental_") ScriptApp.deleteTrigger(t);
  });
  PropertiesService.getScriptProperties().deleteProperty("CALENDAR_SYNC_TOKEN");
  Logger.log("✅ Incremental sync stopped.");
  SpreadsheetApp.getActive().toast("✅ Incremental sync stopped.", "Setup", 5);
}

function setupTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => {
    const fn = t.getHandlerFunction();
    if (fn === "processSheetChanges" || fn === "syncCalendarIncremental_") ScriptApp.deleteTrigger(t);
  });
  PropertiesService.getScriptProperties().deleteProperty("CALENDAR_SYNC_TOKEN");
  ScriptApp.newTrigger("processSheetChanges").forSpreadsheet(SpreadsheetApp.getActive()).onEdit().create();
  ScriptApp.newTrigger("syncCalendarIncremental_").timeBased().everyMinutes(5).create();
  syncCalendarIncremental_();
  const msg = "✅ Triggers installed!\n\n• onEdit → processSheetChanges\n• Every 5 min → syncCalendarIncremental_\n\nIf you see a Calendar error go to:\nServices (+) → Google Calendar API → Add";
  Logger.log(msg);
  SpreadsheetApp.getActive().toast(msg, "Setup", 10);
}

function setupOnOpenSync() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "onOpenSync_") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("onOpenSync_").forSpreadsheet(SpreadsheetApp.getActive()).onOpen().create();
  Logger.log("✅ Sync on sheet open enabled.");
  SpreadsheetApp.getActive().toast("✅ Sync on sheet open enabled.", "Setup", 5);
}

function onOpenSync_() {
  try { syncCalendarIncremental_(); } catch (e) { Logger.log("onOpen sync error: " + e.message); }
}

function validateSetup() {
  const lines = [];

  Logger.log("[validateSetup] START");

  Logger.log("[validateSetup] checking calendar...");
  try {
    getCalendarOrThrow_();
    lines.push("✅ Calendar \"" + CALENDAR_NAME + "\" found");
    Logger.log("[validateSetup] calendar OK");
  } catch (e) {
    lines.push("❌ Calendar: " + e.message);
    Logger.log("[validateSetup] calendar ERROR: " + e.message);
  }

  Logger.log("[validateSetup] checking sheets...");
  const ss = SpreadsheetApp.getActive();
  [APPOINTMENTS_SHEET, CLIENTS_SHEET, SERVICES_SHEET, SUBSCRIPTIONS_SHEET].forEach(name => {
    const found = !!ss.getSheetByName(name);
    lines.push(found ? "✅ Sheet \"" + name + "\" found" : "❌ Sheet \"" + name + "\" MISSING");
    Logger.log("[validateSetup] sheet \"" + name + "\": " + (found ? "OK" : "MISSING"));
  });

  Logger.log("[validateSetup] checking script properties...");
  const props  = PropertiesService.getScriptProperties();
  const token  = props.getProperty("TELEGRAM_BOT_TOKEN");
  const chatId = props.getProperty("TELEGRAM_CHAT_ID");
  lines.push(token  ? "✅ TELEGRAM_BOT_TOKEN set" : "❌ TELEGRAM_BOT_TOKEN not set");
  lines.push(chatId ? "✅ TELEGRAM_CHAT_ID set"   : "❌ TELEGRAM_CHAT_ID not set");
  Logger.log("[validateSetup] TELEGRAM_BOT_TOKEN: " + (token ? "set" : "NOT SET"));
  Logger.log("[validateSetup] TELEGRAM_CHAT_ID: " + (chatId ? "set" : "NOT SET"));

  Logger.log("[validateSetup] checking project triggers...");
  const triggers = ScriptApp.getProjectTriggers().map(t => t.getHandlerFunction());
  Logger.log("[validateSetup] triggers found: " + JSON.stringify(triggers));
  lines.push(triggers.includes("processSheetChanges")      ? "✅ onEdit trigger active"             : "⚠️ onEdit trigger missing — run 🛠️ Setup All Triggers");
  lines.push(triggers.includes("syncCalendarIncremental_") ? "✅ 5-min incremental sync active"     : "⚠️ 5-min sync missing — run ⚡ Setup Incremental Sync");
  lines.push(triggers.includes("sendDailyNotification")    ? "✅ Daily notification trigger active" : "⚠️ Daily notification missing — run setupNotificationTrigger()");
  lines.push(triggers.includes("onOpenSync_")              ? "✅ onOpen sync active"                : "ℹ️ onOpen sync not set (optional — run 📲 Setup Sync on Sheet Open)");

  Logger.log("[validateSetup] checking sync token...");
  const syncToken = props.getProperty("CALENDAR_SYNC_TOKEN");
  lines.push(syncToken ? "✅ Calendar sync token present" : "ℹ️ No sync token yet (runs full sync on first trigger fire)");
  Logger.log("[validateSetup] sync token: " + (syncToken ? "present" : "absent"));

  Logger.log("[validateSetup] all checks done");
  Logger.log("[validateSetup] DONE\n" + lines.join("\n"));
  // Use toast instead of alert so it works both from the editor and the sheet menu
  SpreadsheetApp.getActive().toast(lines.join("\n"), "Setup validation", 15);
}

function syncCalendarIncremental_() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(0)) { Logger.log("Sync already running, skipping."); return; }
  try {
    const props = PropertiesService.getScriptProperties();
    const cal   = getCalendarOrThrow_();
    const calId = cal.getId();
    const tz    = cal.getTimeZone();

    let syncToken = props.getProperty("CALENDAR_SYNC_TOKEN");
    let items = [], newToken = null;

    // Fast path: incremental sync using stored token
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
        // 410 = sync token expired → fall through to full sync
        syncToken = null;
        props.deleteProperty("CALENDAR_SYNC_TOKEN");
      }
    }

    // Full sync: first run or token expired
    if (!syncToken) {
      const now = new Date();
      let resp = Calendar.Events.list(calId, {
        timeMin: startOfDay_(addDays_(now, -DAYS_BACK)).toISOString(),
        timeMax: endOfDay_(addDays_(now,  DAYS_FORWARD)).toISOString(),
        singleEvents: true, maxResults: 2500
      });
      items = resp.items || [];
      while (resp.nextPageToken) {
        resp = Calendar.Events.list(calId, { pageToken: resp.nextPageToken });
        items = items.concat(resp.items || []);
      }
      newToken = resp.nextSyncToken;
    }

    // Nothing changed — save token and exit early
    if (items.length === 0) {
      if (newToken) props.setProperty("CALENDAR_SYNC_TOKEN", newToken);
      return;
    }

    // Something changed — build full context (reads headers + all indexes)
    const ctx = prepareContext_();
    Object.assign(ctx, {
      startDate: startOfDay_(addDays_(new Date(), -DAYS_BACK)),
      endDate:   endOfDay_(addDays_(new Date(),  DAYS_FORWARD))
    });
    const ac    = ctx.apptCols;
    const ncols = ctx.appointmentsSheet.getLastColumn();
    const today = startOfDay_(new Date());

    const newRows    = [];
    let hasChanges   = false;
    const newEventIds = new Set();

    for (const item of items) {
      const eventId = item.iCalUID ? String(item.iCalUID) : String(item.id);

      // Deleted event
      if (item.status === "cancelled") {
        const entry = ctx.apptIndex.get(eventId);
        if (!entry || entry === -1) continue;
        const s = String(entry.data[ac["Status"] - 1] || "");
        if (s !== "Paid" && s !== "No Show" && s !== "Cancelled") {
          ctx.appointmentsSheet.getRange(entry.row, ac["Status"]).setValue("Cancelled");
          ctx.appointmentsSheet.getRange(entry.row, ac["Late"]).setValue(false);
          hasChanges = true;
        }
        continue;
      }

      const parsed = parseEventTitle_(item.summary || "");
      if (!parsed) continue;

      const clientId   = getOrCreateClientId_(ctx.clientsSheet, ctx.clientsIndex, parsed.clientName);
      const hasCredits = hasActiveCredits_(ctx.subsIndex, parsed.clientName, clientId);
      const start      = new Date(item.start.dateTime || item.start.date);
      const isAllDay   = !item.start.dateTime;
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

      const isFuture      = new Date(dateCell) >= today;
      const initialStatus = payment === "Subscription" ? "Paid" : (isFuture ? "Upcoming" : "Not Paid");
      const iExisting     = ctx.apptIndex.get(eventId);

      if (!iExisting) {
        // New appointment
        const rowArr = new Array(ncols).fill("");
        rowArr[ac["Date"] - 1]       = dateCell;
        rowArr[ac["Time"] - 1]       = timeCell;
        rowArr[ac["Price"] - 1]      = price;
        rowArr[ac["Payment"] - 1]    = payment;
        rowArr[ac["Status"] - 1]     = initialStatus;
        rowArr[ac["Tips"] - 1]       = "";
        rowArr[ac["Late"] - 1]       = false;
        rowArr[ac["Notes"] - 1]      = item.description || "";
        rowArr[ac["Service"] - 1]    = serviceToWrite;
        rowArr[ac["ClientID"] - 1]   = clientId;
        rowArr[ac["EventID"] - 1]    = eventId;
        rowArr[ac["Cached Name"] - 1] = parsed.clientName;
        newRows.push(rowArr);

        if (isSubSale) createSubscriptionEntry_(ctx, parsed.clientName, clientId, dateCell);
        ctx.apptIndex.set(eventId, -1);
        newEventIds.add(eventId);
        hasChanges = true;

      } else if (iExisting !== -1) {
        // Updated appointment
        const existingRow = iExisting.row;
        const rowVals     = iExisting.data;
        const oldDate     = new Date(rowVals[ac["Date"] - 1]);
        const oldTimeRaw  = rowVals[ac["Time"] - 1];
        const oldTime     = oldTimeRaw ? hm_(new Date(oldTimeRaw), tz) : "";
        const changed     = ymd !== ymd_(oldDate, tz) ||
                            hm  !== oldTime ||
                            (item.description || "") !== String(rowVals[ac["Notes"] - 1] || "") ||
                            serviceToWrite !== String(rowVals[ac["Service"] - 1] || "") ||
                            parsed.clientName !== String(rowVals[ac["Cached Name"] - 1] || "");
        if (changed) {
          const updatedRow = rowVals.slice();
          updatedRow[ac["Date"] - 1]       = dateCell;
          updatedRow[ac["Time"] - 1]       = timeCell;
          updatedRow[ac["Name"] - 1]       = buildNameFormula_(ctx, existingRow);
          updatedRow[ac["Notes"] - 1]      = item.description || "";
          updatedRow[ac["Service"] - 1]    = serviceToWrite;
          updatedRow[ac["ClientID"] - 1]   = clientId;
          updatedRow[ac["EventID"] - 1]    = eventId;
          updatedRow[ac["Cached Name"] - 1] = parsed.clientName;
          ctx.appointmentsSheet.getRange(existingRow, 1, 1, updatedRow.length).setValues([updatedRow]);
        }
      }
    }

    // Write new rows in one batch
    if (newRows.length > 0) {
      const startRow = ctx.appointmentsSheet.getLastRow() + 1;
      for (let i = 0; i < newRows.length; i++) {
        newRows[i][ac["Name"] - 1] = buildNameFormula_(ctx, startRow + i);
      }
      ctx.appointmentsSheet.getRange(startRow, 1, newRows.length, ncols).setValues(newRows);
    }

    updateUpcomingToNotPaid_(ctx);
    const apptLastRow = ctx.appointmentsSheet.getLastRow();
    const apptData    = loadApptData_(ctx.appointmentsSheet, apptLastRow);
    updateConsecutivePaidCounts_(ctx, apptData);
    updateNoShowLateCounts_(ctx, apptData);
    updateClientStats_(ctx, apptData);
    sortAndHideAppointments_(ctx);

    if (newToken) props.setProperty("CALENDAR_SYNC_TOKEN", newToken);
    if (hasChanges) sendSyncNotification_(ctx, newEventIds);

  } catch (e) {
    Logger.log("Incremental sync error: " + e.message);
  } finally {
    lock.releaseLock();
  }
}

/***************
 * CLEANUP DUPLICATES
 ***************/

function cleanupDuplicates() {
  const ss      = SpreadsheetApp.getActive();
  const sheet   = ss.getSheetByName(APPOINTMENTS_SHEET);
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_ROW) { Logger.log("No data found."); return; }

  const ac   = getSheetCols_(sheet);
  const data = sheet.getRange(DATA_ROW, 1, lastRow - DATA_ROW + 1, sheet.getLastColumn()).getValues();

  const groups = new Map();
  for (let i = 0; i < data.length; i++) {
    const row     = data[i];
    const dateStr = row[ac["Date"] - 1]       ? new Date(row[ac["Date"] - 1]).toDateString() : "";
    const timeStr = row[ac["Time"] - 1]       ? String(row[ac["Time"] - 1]) : "";
    const name    = String(row[ac["Name"] - 1] || row[ac["Cached Name"] - 1] || "").trim().toLowerCase();
    const payment = String(row[ac["Payment"] - 1] || "").trim();
    const status  = String(row[ac["Status"] - 1]  || "").trim();

    if (!dateStr || !name) continue;

    const key = `${dateStr}|${timeStr}|${name}`;
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push({ sheetRow: i + DATA_ROW, payment, status });
  }

  const toDelete = [];
  for (const rows of groups.values()) {
    if (rows.length <= 1) continue;
    const hasGoodRow = rows.some(r =>
      r.payment !== "" || (r.status !== "Not Paid" && r.status !== "Upcoming")
    );
    if (!hasGoodRow) continue;
    for (const r of rows) {
      if (r.payment === "" && (r.status === "Not Paid" || r.status === "Upcoming")) {
        toDelete.push(r.sheetRow);
      }
    }
  }

  if (toDelete.length === 0) {
    Logger.log("✅ No duplicates found — sheet looks clean!");
    SpreadsheetApp.getActive().toast("✅ No duplicates found — sheet looks clean!", "Cleanup", 5);
    return;
  }

  toDelete.sort((a, b) => b - a);
  for (const rowIdx of toDelete) sheet.deleteRow(rowIdx);
  const dupMsg = `✅ Removed ${toDelete.length} duplicate rows. Now run: ✂️ Barber Tools → Setup Incremental Sync`;
  Logger.log(dupMsg);
  SpreadsheetApp.getActive().toast(dupMsg, "Cleanup", 8);
}

function doGet(e) {
  try {
    syncCalendarIncremental_();
    return ContentService.createTextOutput("✅ Sync complete!").setMimeType(ContentService.MimeType.TEXT);
  } catch (err) {
    return ContentService.createTextOutput("❌ Error: " + err.message).setMimeType(ContentService.MimeType.TEXT);
  }
}

// ── FORMATTING ────────────────────────────────────────────────────────────────

function formatSpreadsheet() {
  const COLORS = {
    bg:        '#FAF7F2',
    surface:   '#F2EDE5',
    headerBg:  '#E8DDD0',
    text:      '#2C2017',
    textMuted: '#8C7B6B',
    accent: {
      blue:   '#4A7CA7',
      green:  '#5A8A5A',
      red:    '#B85050',
      yellow: '#B8920A',
      purple: '#8A6A9A',
      orange: '#C47040',
    },
    tint: {
      green:  '#E8F2E8',
      red:    '#F5E4E4',
      yellow: '#F5EDD0',
      purple: '#EDE8F2',
      blue:   '#E4EEF5',
      orange: '#F5EAE0',
    },
  };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  formatAppointments_(ss.getSheetByName('Appointments'), COLORS);
  formatClients_(ss.getSheetByName('Clients'), COLORS);
  formatServices_(ss.getSheetByName('Services'), COLORS);
  formatSubscriptions_(ss.getSheetByName('Subscriptions'), COLORS);
  formatDashboard_(ss.getSheetByName('Dashboard'), COLORS);

  Logger.log('Flushing...');
  SpreadsheetApp.flush();
  Logger.log('✅ Theme applied successfully.');
}

function applyBaseTheme_(sheet, COLORS) {
  if (!sheet) return [];

  Logger.log('Formatting: ' + sheet.getName());

  const maxRows   = sheet.getMaxRows();
  const maxCols   = Math.min(Math.max(sheet.getLastColumn(), 1), sheet.getMaxColumns());
  const styledRows = Math.min(Math.max(sheet.getLastRow(), 1) + 10, maxRows);

  sheet.clearConditionalFormatRules();
  sheet.showColumns(1, maxCols);

  sheet.getRange(1, 1, styledRows, maxCols)
    .setBackground(COLORS.bg)
    .setFontColor(COLORS.text)
    .setFontWeight('normal');

  sheet.getRange(HEADER_ROW, 1, 1, maxCols)
    .setBackground(COLORS.headerBg)
    .setFontColor(COLORS.text)
    .setFontWeight('bold');

  sheet.setFrozenRows(HEADER_ROW);
  sheet.setRowHeight(1, 20);
  sheet.setRowHeight(HEADER_ROW, 320);
  if (maxRows > HEADER_ROW) {
    sheet.setRowHeightsForced(DATA_ROW, maxRows - DATA_ROW + 1, 280);
  }

  const altRule = SpreadsheetApp.newConditionalFormatRule()
    .withCriteria(SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA, ['=MOD(ROW(),2)=0'])
    .setBackground(COLORS.surface)
    .setRanges([sheet.getRange(DATA_ROW, 1, Math.max(maxRows - DATA_ROW + 1, 1), maxCols)])
    .build();

  return [altRule];
}

function formatAppointments_(sheet, COLORS) {
  if (!sheet) return;

  const cols = getSheetCols_(sheet);
  const baseRules = applyBaseTheme_(sheet, COLORS);

  // Column widths by header name — order-independent
  sheet.setColumnWidth(1, 20);
  const widthMap = {
    'Date': 140, 'Time': 110, 'Name': 290, 'Price': 110,
    'Payment': 185, 'Status': 155, 'Tips': 110, 'Late': 90,
    'Notes': 240, 'Service': 210, 'ClientID': 80, 'EventID': 80, 'Cached Name': 80
  };
  Object.entries(widthMap).forEach(([name, w]) => {
    if (cols[name]) sheet.setColumnWidth(cols[name], w);
  });

  // Hide internal columns by name
  ['ClientID', 'EventID', 'Cached Name'].forEach(name => {
    if (cols[name]) sheet.hideColumns(cols[name], 1);
  });

  const maxRows  = sheet.getMaxRows();
  const dataRows = Math.max(maxRows - DATA_ROW + 1, 1);
  const lastCol  = sheet.getLastColumn();
  const dataRange = sheet.getRange(DATA_ROW, 2, dataRows, lastCol - 1);

  dataRange
    .setFontSize(26)
    .setVerticalAlignment('middle')
    .setFontWeight('normal')
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

  sheet.getRange(HEADER_ROW, 2, 1, lastCol - 1).setFontSize(18);

  if (cols['Time'])    sheet.getRange(DATA_ROW, cols['Time'],    dataRows, 1).setFontSize(30).setFontWeight('bold');
  if (cols['Name'])    sheet.getRange(DATA_ROW, cols['Name'],    dataRows, 1).setFontSize(30).setFontWeight('bold').setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  if (cols['Date'])    sheet.getRange(DATA_ROW, cols['Date'],    dataRows, 1).setFontSize(20);
  if (cols['Notes'])   sheet.getRange(DATA_ROW, cols['Notes'],   dataRows, 1).setFontSize(22).setFontColor(COLORS.textMuted);
  if (cols['Service']) sheet.getRange(DATA_ROW, cols['Service'], dataRows, 1).setFontSize(22).setFontColor(COLORS.textMuted);

  ['Date', 'Time', 'Price', 'Payment', 'Status', 'Tips', 'Late'].forEach(name => {
    if (cols[name]) sheet.getRange(DATA_ROW, cols[name], dataRows, 1).setHorizontalAlignment('center');
  });

  // Build $COL letter references for conditional format formulas
  const L = name => cols[name] ? '$' + colLetter_(cols[name]) : null;
  const ruleDefs = [
    { col: L('Late'),    formula: `=${L('Late')}${DATA_ROW}=TRUE`,              bg: COLORS.tint.orange, fg: COLORS.accent.orange },
    { col: L('Status'),  formula: `=${L('Status')}${DATA_ROW}="No Show"`,       bg: COLORS.tint.red,    fg: COLORS.accent.red    },
    { col: L('Status'),  formula: `=${L('Status')}${DATA_ROW}="Cancelled"`,     bg: COLORS.bg,          fg: COLORS.textMuted     },
    { col: L('Payment'), formula: `=${L('Payment')}${DATA_ROW}="Subscription"`, bg: COLORS.tint.purple, fg: COLORS.accent.purple },
    { col: L('Status'),  formula: `=${L('Status')}${DATA_ROW}="Paid"`,          bg: COLORS.tint.green,  fg: COLORS.accent.green  },
    { col: L('Status'),  formula: `=${L('Status')}${DATA_ROW}="Upcoming"`,      bg: COLORS.tint.blue,   fg: COLORS.accent.blue   },
    { col: L('Status'),  formula: `=${L('Status')}${DATA_ROW}="Not Paid"`,      bg: COLORS.tint.yellow, fg: COLORS.accent.yellow },
  ];

  const rules = ruleDefs
    .filter(r => r.col)
    .map(r =>
      SpreadsheetApp.newConditionalFormatRule()
        .withCriteria(SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA, [r.formula])
        .setBackground(r.bg).setFontColor(r.fg).setRanges([dataRange]).build()
    );

  sheet.setConditionalFormatRules([...rules, ...baseRules]);
  sheet.setTabColor(COLORS.accent.green);
}

function formatClients_(sheet, COLORS) {
  if (!sheet) return;

  const cols = getSheetCols_(sheet);
  const baseRules = applyBaseTheme_(sheet, COLORS);

  // Column widths by header name — order-independent
  sheet.setColumnWidth(1, 20);
  const widthMap = {
    'Name': 280, 'Favourite Service': 240, 'Last Visit': 170, 'Social Media': 100,
    'Notes': 310, 'No Show': 140, 'Late': 120, 'Referral': 100,
    'Total Visits': 100, 'Total Tips': 100, 'Total Spent': 100,
    'ClientID': 100, 'First Visit': 100, 'Do Not Cut': 130, 'Consecutive Paid': 130, 'VIP': 100
  };
  Object.entries(widthMap).forEach(([name, w]) => {
    if (cols[name]) sheet.setColumnWidth(cols[name], w);
  });

  // Hide internal columns by name
  ['Social Media', 'Referral', 'Total Visits', 'Total Tips', 'Total Spent', 'ClientID', 'First Visit'].forEach(name => {
    if (cols[name]) sheet.hideColumns(cols[name], 1);
  });

  const maxRows  = sheet.getMaxRows();
  const dataRows = Math.max(maxRows - DATA_ROW + 1, 1);
  const lastCol  = sheet.getLastColumn();
  const dataRange = sheet.getRange(DATA_ROW, 2, dataRows, lastCol - 1);

  dataRange
    .setFontSize(26)
    .setVerticalAlignment('middle')
    .setFontWeight('normal')
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  sheet.getRange(HEADER_ROW, 2, 1, lastCol - 1).setFontSize(18);

  if (cols['Name'])             sheet.getRange(DATA_ROW, cols['Name'],             dataRows, 1).setFontSize(30).setFontWeight('bold').setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  if (cols['Favourite Service']) sheet.getRange(DATA_ROW, cols['Favourite Service'], dataRows, 1).setFontSize(22);
  if (cols['Notes'])            sheet.getRange(DATA_ROW, cols['Notes'],            dataRows, 1).setFontSize(22).setFontColor(COLORS.textMuted);

  ['Last Visit', 'No Show', 'Late', 'Do Not Cut', 'Consecutive Paid', 'VIP'].forEach(name => {
    if (cols[name]) sheet.getRange(DATA_ROW, cols[name], dataRows, 1).setHorizontalAlignment('center');
  });

  const L = name => cols[name] ? '$' + colLetter_(cols[name]) : null;
  const noShowL = L('No Show'), lateL = L('Late');
  const ruleDefs = [
    { col: L('Do Not Cut'), formula: `=${L('Do Not Cut')}${DATA_ROW}=TRUE`,                         bg: COLORS.tint.red,    fg: COLORS.accent.red    },
    { col: noShowL && lateL, formula: `=${noShowL}${DATA_ROW}+${lateL}${DATA_ROW}>=3`,              bg: COLORS.tint.orange, fg: COLORS.accent.orange },
    { col: L('VIP'),        formula: `=${L('VIP')}${DATA_ROW}=TRUE`,                                bg: COLORS.tint.yellow, fg: COLORS.accent.yellow },
  ];

  const rules = ruleDefs
    .filter(r => r.col)
    .map(r =>
      SpreadsheetApp.newConditionalFormatRule()
        .withCriteria(SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA, [r.formula])
        .setBackground(r.bg).setFontColor(r.fg).setRanges([dataRange]).build()
    );

  sheet.setConditionalFormatRules([...rules, ...baseRules]);
  sheet.setTabColor(COLORS.accent.purple);
}

function formatServices_(sheet, COLORS) {
  if (!sheet) return;
  const cols = getSheetCols_(sheet);
  const baseRules = applyBaseTheme_(sheet, COLORS);
  sheet.setColumnWidth(1, 20);
  if (cols['Service']) sheet.setColumnWidth(cols['Service'], 150);
  if (cols['Price'])   sheet.setColumnWidth(cols['Price'],   90);
  sheet.setConditionalFormatRules(baseRules);
  sheet.setTabColor(COLORS.textMuted);
}

function formatSubscriptions_(sheet, COLORS) {
  if (!sheet) return;
  const baseRules = applyBaseTheme_(sheet, COLORS);
  sheet.setConditionalFormatRules(baseRules);
  sheet.setTabColor(COLORS.accent.orange);
}

function formatDashboard_(sheet, COLORS) {
  if (!sheet) return;

  const baseRules  = applyBaseTheme_(sheet, COLORS);
  const styledRows = Math.min(Math.max(sheet.getLastRow(), 1) + 5, 50);
  const styledCols = Math.min(Math.max(sheet.getLastColumn(), 1), 15);

  sheet.getRange(1, 2, styledRows, 1).setFontColor(COLORS.textMuted);
  [3, 8, 9, 10].forEach(col => {
    if (col <= styledCols) sheet.getRange(1, col, styledRows, 1).setFontWeight('bold').setFontColor(COLORS.text);
  });
  if (styledCols >= 10) sheet.getRange('J2').setFontColor(COLORS.accent.blue).setFontWeight('bold');
  if (styledCols >= 3) {
    sheet.getRange(3, 1, 1, styledCols).setBackground(COLORS.tint.blue).setFontWeight('bold');
  }

  sheet.setConditionalFormatRules(baseRules);
  sheet.setTabColor(COLORS.accent.blue);
}
