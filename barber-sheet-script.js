/***************
 * CONFIGURATION
 * Change these values to adapt the script to your setup.
 ***************/

// Name of the Google Calendar to sync from. Must match exactly.
const CALENDAR_NAME = "Barber Appointments";

// Names of the four required sheets inside the spreadsheet.
// Change these if your sheet tabs have different names.
const APPOINTMENTS_SHEET = "Appointments";
const CLIENTS_SHEET = "Clients";
const SERVICES_SHEET = "Services";
const SUBSCRIPTIONS_SHEET = "Subscriptions";

// How many days back/forward to include in a full sync.
const DAYS_BACK = 14;
const DAYS_FORWARD = 14;

// Appointments older than this many days are hidden (not deleted).
const HIDE_OLDER_THAN_DAYS = 30;

// Notification channel. Currently only "telegram" is supported.
// Requires TELEGRAM_BOT_TOKEN and TELEGRAM_CHAT_ID in Script Properties.
const NOTIFICATION_MODE = "telegram";

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
    .addItem("🔄 Sync Now", "syncCalendarIncremental_")
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
  if (sheetName === APPOINTMENTS_SHEET && range.getRow() > 1) {
    const row = range.getRow();

    // A. Payment type changed (col E)
    if (range.getColumn() === 5) {
      const paymentVal = range.getValue();
      const statusRange = sheet.getRange(row, 6);
      const currentStatus = statusRange.getValue();

      // Never override a manually set No Show or Cancelled
      if (currentStatus !== "No Show" && currentStatus !== "Cancelled") {
        if (paymentVal === "Cash" || paymentVal === "Tikkie" || paymentVal === "Subscription" || paymentVal === "Free") {
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

      if (paymentVal === "Subscription" || paymentVal === "Free") {
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

    // C. Recalculate all client stats immediately when any stat-affecting column changes.
    // Covers: D=Price, E=Payment, F=Status, G=Tips, H=Late, K=ClientID
    // Runs after A and B so their sheet writes are already committed before we re-read.
    const col = range.getColumn();
    if (col === 4 || col === 5 || col === 6 || col === 7 || col === 8 || col === 11) {
      const minCtx = {
        appointmentsSheet: sheet,
        clientsSheet: sheet.getParent().getSheetByName(CLIENTS_SHEET)
      };
      updateClientStats_(minCtx);
      updateNoShowLateCounts_(minCtx);
      updateConsecutivePaidCounts_(minCtx);
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
  const apptLastRow = ctx.appointmentsSheet.getLastRow();
  const apptData = apptLastRow >= 2 ? ctx.appointmentsSheet.getRange(2, 1, apptLastRow - 1, 13).getValues() : [];
  updateConsecutivePaidCounts_(ctx, apptData);
  updateNoShowLateCounts_(ctx, apptData);
  updateClientStats_(ctx, apptData);
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
  const apptLastRowY = ctx.appointmentsSheet.getLastRow();
  const apptDataY = apptLastRowY >= 2 ? ctx.appointmentsSheet.getRange(2, 1, apptLastRowY - 1, 13).getValues() : [];
  updateConsecutivePaidCounts_(ctx, apptDataY);
  updateNoShowLateCounts_(ctx, apptDataY);
  updateClientStats_(ctx, apptDataY);
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

    const existing = ctx.apptIndex.get(eventId);
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

    if (!existing) {
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
      if (existing === -1) continue;

      const existingRow = existing.row;
      const rowVals = existing.data;

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
  ctx.apptIndex.forEach((entry, eId) => {
    if (entry === -1) return;
    if (validEventIds.has(eId)) return;

    const rowIdx = entry.row;
    const dateVal = entry.data[0]; // col A
    if (!dateVal) return;
    const rowDate = new Date(dateVal);
    if (rowDate < ctx.startDate || rowDate > ctx.endDate) return;

    const currentStatus = String(entry.data[5] || ""); // col F
    if (currentStatus === "Paid" || currentStatus === "No Show" || currentStatus === "Cancelled") return;

    ctx.appointmentsSheet.getRange(rowIdx, 6).setValue("Cancelled");
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
function updateConsecutivePaidCounts_(ctx, apptData) {
  const clientsSheet = ctx.clientsSheet;
  const apptSheet = ctx.appointmentsSheet;
  const lastRow = clientsSheet.getLastRow();
  if (lastRow < 2) return;

  const clientData = clientsSheet.getRange(2, 1, lastRow - 1, 16).getValues();
  if (!apptData) {
    const apptLastRow = apptSheet.getLastRow();
    apptData = apptLastRow >= 2 ? apptSheet.getRange(2, 1, apptLastRow - 1, 13).getValues() : [];
  }

  const n = clientData.length;
  const oVals = new Array(n);

  for (let i = 0; i < n; i++) {
    const clientId = clientData[i][11]; // col L
    if (!clientId) { oVals[i] = [0]; continue; }

    const clientAppts = [];
    for (const row of apptData) {
      if (String(row[10]) === String(clientId)) {
        clientAppts.push({ date: row[0], status: row[5], payment: row[4], late: row[7] });
      }
    }
    clientAppts.sort((a, b) => new Date(b.date) - new Date(a.date));

    let consecutivePaid = 0;
    for (const appt of clientAppts) {
      if (appt.status === "No Show" || appt.late === true) break;
      const isPaid = appt.payment === "Cash" || appt.payment === "Tikkie" || appt.payment === "Subscription" || appt.payment === "Free" || appt.status.startsWith("Free");
      if (isPaid) consecutivePaid++;
      else break;
    }
    oVals[i] = [consecutivePaid];
  }

  clientsSheet.getRange(2, 15, n, 1).setValues(oVals); // col O
}

/**
 * Update TotalVisits (I), TotalTips (J), TotalSpent (K), LastVisit (C), FirstVisit (M)
 * in Clients sheet. Replaces spreadsheet formulas so data stays correct after
 * historical appointment rows are deleted.
 *
 * "Paid visit" = payment is Cash/Tikkie/Subscription/Free OR status starts with "Free".
 * TotalSpent counts only Cash/Tikkie/Subscription (Free appointments are €0 by definition).
 */
function updateClientStats_(ctx, apptData) {
  const clientsSheet = ctx.clientsSheet;
  const apptSheet = ctx.appointmentsSheet;
  const clientLastRow = clientsSheet.getLastRow();
  if (clientLastRow < 2) return;

  const clientData = clientsSheet.getRange(2, 1, clientLastRow - 1, 12).getValues();
  if (!apptData) {
    const apptLastRow = apptSheet.getLastRow();
    apptData = apptLastRow >= 2 ? apptSheet.getRange(2, 1, apptLastRow - 1, 11).getValues() : [];
  }

  const lastVisits = [];   // col C
  const ijkValues = [];    // cols I, J, K contiguous
  const firstVisits = [];  // col M

  for (let i = 0; i < clientData.length; i++) {
    const clientId = clientData[i][11];
    if (!clientId) {
      lastVisits.push([""]);
      ijkValues.push([0, 0, 0]);
      firstVisits.push([""]);
      continue;
    }

    let totalVisits = 0, totalTips = 0, totalSpent = 0;
    let firstVisit = null, lastVisit = null;

    for (const row of apptData) {
      if (String(row[10]) !== String(clientId)) continue;

      const payment = String(row[4] || "");
      const status  = String(row[5] || "");
      const isPaidVisit = payment === "Cash" || payment === "Tikkie" ||
                          payment === "Subscription" || payment === "Free" ||
                          status.startsWith("Free");
      if (!isPaidVisit) continue;

      const dateVal = row[0];
      if (dateVal) {
        const d = new Date(dateVal);
        if (!firstVisit || d < firstVisit) firstVisit = d;
        if (!lastVisit  || d > lastVisit)  lastVisit  = d;
      }

      totalVisits++;
      totalTips += Number(row[6]) || 0;
      if (payment === "Cash" || payment === "Tikkie" || payment === "Subscription") {
        totalSpent += Number(row[3]) || 0;
      }
    }

    lastVisits.push([lastVisit  || ""]);
    ijkValues.push([totalVisits, totalTips, totalSpent]);
    firstVisits.push([firstVisit || ""]);
  }

  const n = clientData.length;
  clientsSheet.getRange(2, 3,  n, 1).setValues(lastVisits);  // col C — LastVisit
  clientsSheet.getRange(2, 9,  n, 3).setValues(ijkValues);   // cols I, J, K
  clientsSheet.getRange(2, 13, n, 1).setValues(firstVisits); // col M — FirstVisit
}

/**
 * Update NoShow (col F) and Late (col G) counts in Clients sheet.
 * Replaces sheet formulas — counts from appointment history over last 12 months.
 * Call this after every sync instead of maintaining COUNTIFS formulas in the sheet.
 */
function updateNoShowLateCounts_(ctx, apptData) {
  const clientsSheet = ctx.clientsSheet;
  const apptSheet = ctx.appointmentsSheet;
  const clientLastRow = clientsSheet.getLastRow();
  if (clientLastRow < 2) return;

  if (!apptData) {
    const apptLastRow = apptSheet.getLastRow();
    if (apptLastRow < 2) return;
    apptData = apptSheet.getRange(2, 1, apptLastRow - 1, 11).getValues();
  }

  const cutoff = new Date();
  cutoff.setFullYear(cutoff.getFullYear() - 1);

  const clientData = clientsSheet.getRange(2, 1, clientLastRow - 1, 12).getValues();
  const n = clientData.length;
  const fgVals = new Array(n);

  for (let i = 0; i < n; i++) {
    const clientId = clientData[i][11]; // col L
    if (!clientId) { fgVals[i] = [0, 0]; continue; }

    let noShows = 0, lates = 0;
    for (const row of apptData) {
      if (String(row[10]) !== String(clientId)) continue;
      const dateVal = row[0];
      if (!dateVal || new Date(dateVal) < cutoff) continue;
      if (row[5] === "No Show") noShows++;
      if (row[7] === true) lates++;
    }
    fgVals[i] = [noShows, lates];
  }

  clientsSheet.getRange(2, 6, n, 2).setValues(fgVals); // cols F–G
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
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return map;
  const data = sheet.getRange(2, 1, lastRow - 1, 13).getValues();
  for (let i = 0; i < data.length; i++) {
    const eventId = data[i][11]; // col L
    if (eventId) map.set(String(eventId), { row: i + 2, data: data[i] });
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

/**
 * ONE-CLICK SETUP — installs onEdit trigger (processSheetChanges) + 5-min incremental sync.
 * Run this once from the desktop Apps Script editor after pasting the script.
 * Also requires: Services → Google Calendar API → Add (for incremental sync).
 */
function setupTriggers() {
  // Remove existing to avoid duplicates
  ScriptApp.getProjectTriggers().forEach(t => {
    const fn = t.getHandlerFunction();
    if (fn === "processSheetChanges" || fn === "syncCalendarIncremental_") {
      ScriptApp.deleteTrigger(t);
    }
  });
  PropertiesService.getScriptProperties().deleteProperty("CALENDAR_SYNC_TOKEN");

  ScriptApp.newTrigger("processSheetChanges").forSpreadsheet(SpreadsheetApp.getActive()).onEdit().create();
  ScriptApp.newTrigger("syncCalendarIncremental_").timeBased().everyMinutes(5).create();

  syncCalendarIncremental_(); // first run to get initial sync token
  safeAlert_("✅ Triggers installed!\n\n• onEdit → processSheetChanges (Dashboard button + sheet logic)\n• Every 5 min → syncCalendarIncremental_\n\nIf you see a Calendar error go to:\nServices (+) → Google Calendar API → Add");
}

/**
 * Installs an installable onOpen trigger so the sheet syncs automatically when opened.
 * Run this once. This is separate from setupTriggers because it requires explicit authorization.
 */
function setupOnOpenSync() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "onOpenSync_") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("onOpenSync_").forSpreadsheet(SpreadsheetApp.getActive()).onOpen().create();
  safeAlert_("✅ Sync on sheet open enabled.\nThe sheet will now run an incremental sync every time it is opened.");
}

function onOpenSync_() {
  try { syncCalendarIncremental_(); } catch (e) { Logger.log("onOpen sync error: " + e.message); }
}

/**
 * Checks that everything is configured correctly.
 * Run from ✂️ Barber Tools → Validate Setup.
 */
function validateSetup() {
  const lines = [];

  // 1. Calendar
  try {
    getCalendarOrThrow_();
    lines.push("✅ Calendar \"" + CALENDAR_NAME + "\" found");
  } catch (e) {
    lines.push("❌ Calendar: " + e.message);
  }

  // 2. Sheets
  const ss = SpreadsheetApp.getActive();
  [APPOINTMENTS_SHEET, CLIENTS_SHEET, SERVICES_SHEET, SUBSCRIPTIONS_SHEET].forEach(name => {
    lines.push(ss.getSheetByName(name) ? "✅ Sheet \"" + name + "\" found" : "❌ Sheet \"" + name + "\" MISSING");
  });

  // 3. Telegram
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty("TELEGRAM_BOT_TOKEN");
  const chatId = props.getProperty("TELEGRAM_CHAT_ID");
  lines.push(token ? "✅ TELEGRAM_BOT_TOKEN set" : "❌ TELEGRAM_BOT_TOKEN not set");
  lines.push(chatId ? "✅ TELEGRAM_CHAT_ID set" : "❌ TELEGRAM_CHAT_ID not set");

  // 4. Triggers
  const triggers = ScriptApp.getProjectTriggers().map(t => t.getHandlerFunction());
  lines.push(triggers.includes("processSheetChanges") ? "✅ onEdit trigger active" : "⚠️ onEdit trigger missing — run 🛠️ Setup All Triggers");
  lines.push(triggers.includes("syncCalendarIncremental_") ? "✅ 5-min incremental sync active" : "⚠️ 5-min sync missing — run ⚡ Setup Incremental Sync");
  lines.push(triggers.includes("sendDailyNotification") ? "✅ Daily notification trigger active" : "⚠️ Daily notification missing — run setupNotificationTrigger()");
  lines.push(triggers.includes("onOpenSync_") ? "✅ onOpen sync active" : "ℹ️ onOpen sync not set (optional — run 📲 Setup Sync on Sheet Open)");

  // 5. Sync token
  const syncToken = props.getProperty("CALENDAR_SYNC_TOKEN");
  lines.push(syncToken ? "✅ Calendar sync token present" : "ℹ️ No sync token yet (runs full sync on first trigger fire)");

  safeAlert_("Setup validation:\n\n" + lines.join("\n"));
  Logger.log(lines.join("\n"));
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
        const entry = ctx.apptIndex.get(eventId);
        if (!entry || entry === -1) continue;
        const rowIdx = entry.row;
        const s = String(entry.data[5] || ""); // col F = Status
        if (s !== "Paid" && s !== "No Show" && s !== "Cancelled") {
          apptSheet.getRange(rowIdx, 6).setValue("Cancelled");
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
      const iExisting = ctx.apptIndex.get(eventId);

      if (!iExisting) {
        // New appointment
        newRowsABC.push([dateCell, timeCell, ""]);
        newRowsDToM.push([price, payment, initialStatus, "", false,
          item.description || "", serviceToWrite, clientId, eventId, parsed.clientName]);
        if (isSubSale) createSubscriptionEntry_(ctx, parsed.clientName, clientId, dateCell);
        ctx.apptIndex.set(eventId, -1);
        newEventIds.add(eventId);
        hasChanges = true;

      } else if (iExisting !== -1) {
        // Updated appointment — refresh time / notes / service
        const existingRow = iExisting.row;
        const rowVals     = iExisting.data;
        const oldDate     = new Date(rowVals[0]);
        const oldTime     = rowVals[1] ? hm_(new Date(rowVals[1]), tz) : "";
        const changed     = ymd !== ymd_(oldDate, tz) ||
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
    const iLastRow = apptSheet.getLastRow();
    const iApptData = iLastRow >= 2 ? apptSheet.getRange(2, 1, iLastRow - 1, 13).getValues() : [];
    updateConsecutivePaidCounts_(ctx, iApptData);
    updateNoShowLateCounts_(ctx, iApptData);
    updateClientStats_(ctx, iApptData);
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

// ── FORMATTING ────────────────────────────────────────────────────────────────

/**
 * Applies a dark/charcoal visual theme with pastel accents to all sheets.
 * Run from the Apps Script editor dropdown. Safe to re-run at any time.
 */
function formatSpreadsheet() {
  const COLORS = {
    bg:        '#FAF7F2',  // warm cream base
    surface:   '#F2EDE5',  // slightly darker cream (alternating rows)
    headerBg:  '#E8DDD0',  // warm tan header
    text:      '#2C2017',  // deep warm brown
    textMuted: '#8C7B6B',  // muted mid-brown
    accent: {
      blue:   '#4A7CA7',  // steel blue
      green:  '#5A8A5A',  // sage green
      red:    '#B85050',  // terracotta
      yellow: '#B8920A',  // amber
      purple: '#8A6A9A',  // mauve
      orange: '#C47040',  // burnt orange
    },
    tint: {
      green:  '#E8F2E8',  // light sage
      red:    '#F5E4E4',  // light terracotta
      yellow: '#F5EDD0',  // light amber
      purple: '#EDE8F2',  // light mauve
      blue:   '#E4EEF5',  // light steel
      orange: '#F5EAE0',  // light burnt orange
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

/**
 * Sets the dark base on a sheet: clears old rules, dark bg, styled header, freezes row 1.
 * Returns a single alternating-row rule to be appended last (lowest priority) by the caller.
 */
function applyBaseTheme_(sheet, COLORS) {
  if (!sheet) return [];

  Logger.log('Formatting: ' + sheet.getName());

  const maxRows = sheet.getMaxRows();
  const maxCols = Math.min(Math.max(sheet.getLastColumn(), 1), sheet.getMaxColumns());
  // Style only up to the actual sheet boundary — never exceed maxRows or Google Sheets adds rows
  const styledRows = Math.min(Math.max(sheet.getLastRow(), 1) + 10, maxRows);

  // Clear existing conditional format rules (idempotency)
  sheet.clearConditionalFormatRules();

  // Unhide all columns so re-runs don't stack hidden columns
  sheet.showColumns(1, maxCols);

  // Base — skip setFontFamily/setFontSize, they are very slow on large ranges
  sheet.getRange(1, 1, styledRows, maxCols)
    .setBackground(COLORS.bg)
    .setFontColor(COLORS.text)
    .setFontWeight('normal');

  // Header row: warm tan bg, dark text bold
  sheet.getRange(1, 1, 1, maxCols)
    .setBackground(COLORS.headerBg)
    .setFontColor(COLORS.text)
    .setFontWeight('bold');

  sheet.setFrozenRows(1);

  // Taller rows for better tap targets on mobile (26px)
  if (maxRows > 1) {
    sheet.setRowHeightsForced(2, maxRows - 1, 26);
  }

  // Alternating row rule — must go last (lowest priority) so status colors win
  // Use maxRows (not a hardcoded large number) to avoid auto-expanding the sheet
  const altRule = SpreadsheetApp.newConditionalFormatRule()
    .withCriteria(SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA, ['=MOD(ROW(),2)=0'])
    .setBackground(COLORS.surface)
    .setRanges([sheet.getRange(2, 1, Math.max(maxRows - 1, 1), maxCols)])
    .build();

  return [altRule];
}

/**
 * Formats the Appointments sheet.
 * Columns: A=Date B=Time C=Name D=Price E=Payment F=Status G=Tips H=Late I=Notes J=Service K=ClientID L=EventID M=CachedName
 */
function formatAppointments_(sheet, COLORS) {
  if (!sheet) return;

  const baseRules = applyBaseTheme_(sheet, COLORS);

  // Portrait-optimised widths: A–D (58+58+120+58=294px) fit in ~360px viewport,
  // so Payment (E) is just at the edge — one small scroll reveals the action column.
  const widths = [58, 58, 120, 58, 90, 90, 55, 45, 140, 100, 80, 80, 80];
  widths.forEach((w, i) => { if (w > 0) sheet.setColumnWidth(i + 1, w); });

  // Hide internal ID/lookup columns: K=ClientID L=EventID M=CachedName
  sheet.hideColumns(11, 3);

  const dataRange = sheet.getRange(2, 1, Math.max(sheet.getMaxRows() - 1, 1), 13);

  // Status/payment rules in priority order (index 0 = highest priority)
  // Late checkbox (H) wins over everything — makes overdue/late appointments immediately visible
  const ruleDefs = [
    { formula: '=$H2=TRUE',              bg: COLORS.tint.orange, fg: COLORS.accent.orange },
    { formula: '=$F2="No Show"',         bg: COLORS.tint.red,    fg: COLORS.accent.red    },
    { formula: '=$F2="Cancelled"',       bg: COLORS.bg,          fg: COLORS.textMuted     },
    { formula: '=$E2="Subscription"',    bg: COLORS.tint.purple, fg: COLORS.accent.purple },
    { formula: '=$F2="Paid"',            bg: COLORS.tint.green,  fg: COLORS.accent.green  },
    { formula: '=$F2="Upcoming"',        bg: COLORS.tint.blue,   fg: COLORS.accent.blue   },
    { formula: '=$F2="Not Paid"',        bg: COLORS.tint.yellow, fg: COLORS.accent.yellow },
  ];

  const rules = ruleDefs.map(r =>
    SpreadsheetApp.newConditionalFormatRule()
      .withCriteria(SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA, [r.formula])
      .setBackground(r.bg)
      .setFontColor(r.fg)
      .setRanges([dataRange])
      .build()
  );

  sheet.setConditionalFormatRules([...rules, ...baseRules]);
  sheet.setTabColor(COLORS.accent.green);  // sage green tab
}

/**
 * Formats the Clients sheet.
 * Columns: A=Name B=FavService C=LastVisit D=SocialMedia E=Notes F=NoShow(12m) G=Late(12m)
 *          H=Referral I=TotalVisits J=TotalTips K=TotalSpent L=ClientID M=FirstVisit
 *          N=DoNotCut O=ConsecutivePaid P=VIP
 */
function formatClients_(sheet, COLORS) {
  if (!sheet) return;

  const baseRules = applyBaseTheme_(sheet, COLORS);

  // A=Name, B=FavService, C=LastVisit, D=SocialMedia, E=Notes, F=NoShow, G=Late,
  // H=Referral, I=TotalVisits, J=TotalTips, K=TotalSpent, L=ClientID, M=FirstVisit,
  // N=DoNotCut, O=ConsecutivePaid, P=VIP
  const widths = [130, 105, 90, 100, 140, 75, 65, 90, 70, 65, 75, 70, 80, 70, 80, 50];
  widths.forEach((w, i) => { if (w > 0) sheet.setColumnWidth(i + 1, w); });

  // Hide statistical/rarely-checked columns on mobile:
  // D=SocialMedia(4), H=Referral(8), I=TotalVisits(9), J=TotalTips(10), K=TotalSpent(11),
  // L=ClientID(12), M=FirstVisit(13)
  sheet.hideColumns(4, 1);   // D SocialMedia
  sheet.hideColumns(8, 6);   // H–M (Referral, TotalVisits, TotalTips, TotalSpent, ClientID, FirstVisit)

  const dataRange = sheet.getRange(2, 1, Math.max(sheet.getMaxRows() - 1, 1), 16);

  // DoNotCut has highest urgency — must be immediately visible
  const ruleDefs = [
    { formula: '=$N2=TRUE',   bg: COLORS.tint.red,    fg: COLORS.accent.red    },  // Do Not Cut
    { formula: '=$F2+$G2>=3', bg: COLORS.tint.orange, fg: COLORS.accent.orange },  // NoShow+Late(12m) ≥ 3
    { formula: '=$P2=TRUE',   bg: COLORS.tint.yellow, fg: COLORS.accent.yellow },  // VIP
  ];

  const rules = ruleDefs.map(r =>
    SpreadsheetApp.newConditionalFormatRule()
      .withCriteria(SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA, [r.formula])
      .setBackground(r.bg)
      .setFontColor(r.fg)
      .setRanges([dataRange])
      .build()
  );

  sheet.setConditionalFormatRules([...rules, ...baseRules]);
  sheet.setTabColor(COLORS.accent.purple);  // mauve tab
}

/** Formats the Services sheet. */
function formatServices_(sheet, COLORS) {
  if (!sheet) return;
  const baseRules = applyBaseTheme_(sheet, COLORS);
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 90);
  sheet.setConditionalFormatRules(baseRules);
  sheet.setTabColor(COLORS.textMuted);  // muted brown tab — de-emphasised
}

/** Formats the Subscriptions sheet. */
function formatSubscriptions_(sheet, COLORS) {
  if (!sheet) return;
  const baseRules = applyBaseTheme_(sheet, COLORS);
  sheet.setConditionalFormatRules(baseRules);
  sheet.setTabColor(COLORS.accent.orange);  // burnt orange tab — distinct from Clients mauve
}

/**
 * Formats the Dashboard sheet.
 * Sync checkbox is at C3 — clearFormats is avoided to preserve data validation.
 */
function formatDashboard_(sheet, COLORS) {
  if (!sheet) return;

  const baseRules = applyBaseTheme_(sheet, COLORS);

  const styledRows = Math.min(Math.max(sheet.getLastRow(), 1) + 5, 50);
  const styledCols = Math.min(Math.max(sheet.getLastColumn(), 1), 15);

  // Label column (B) dimmed; value columns bold
  sheet.getRange(1, 2, styledRows, 1).setFontColor(COLORS.textMuted);
  [3, 8, 9, 10].forEach(col => {
    if (col <= styledCols) sheet.getRange(1, col, styledRows, 1).setFontWeight('bold').setFontColor(COLORS.text);
  });
  if (styledCols >= 10) sheet.getRange('J2').setFontColor(COLORS.accent.blue).setFontWeight('bold');

  // Row 3 = sync checkbox row (C3). Give it a warm tint to stand out as the tap target.
  if (styledCols >= 3) {
    sheet.getRange(3, 1, 1, styledCols)
      .setBackground(COLORS.tint.blue)
      .setFontWeight('bold');
  }

  sheet.setConditionalFormatRules(baseRules);
  sheet.setTabColor(COLORS.accent.blue);  // steel blue tab
}
