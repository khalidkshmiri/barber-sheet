/***************
 * CONFIGURATION
 * Change these values to adapt the script to your setup.
 ***************/

// Name of the Google Calendar to sync from. Must match exactly.
const CALENDAR_NAME = "Barber Appointments";

// Names of the four required sheets inside the spreadsheet.
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
const NOTIFICATION_MODE = "telegram";

// Layout constants — tables begin at B2 (row 1 and col A are spacers).
const HEADER_ROW = 2;  // row that holds column headers
const DATA_ROW   = 3;  // first row that holds actual data

// Appointments columns (1-based sheet column numbers, col A = spacer):
// B=2 Date  C=3 Time  D=4 Name(formula)  E=5 Price  F=6 Payment
// G=7 Status  H=8 Tips  I=9 Late  J=10 Notes  K=11 Service
// L=12 ClientID  M=13 EventID  N=14 CachedName

// Clients columns (1-based):
// B=2 Name  C=3 FavService  D=4 LastVisit  E=5 SocialMedia  F=6 Notes
// G=7 NoShow(12m)  H=8 Late(12m)  I=9 Referral  J=10 TotalVisits
// K=11 TotalTips  L=12 TotalSpent  M=13 ClientID  N=14 FirstVisit
// O=15 DoNotCut  P=16 ConsecutivePaid (auto)  Q=17 VIP

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
  if (sheetName === APPOINTMENTS_SHEET && range.getRow() > HEADER_ROW) {
    const row = range.getRow();

    // A. Payment type changed (col F=6)
    if (range.getColumn() === 6) {
      const paymentVal = range.getValue();
      const statusRange = sheet.getRange(row, 7); // col G = Status
      const currentStatus = statusRange.getValue();

      // Never override a manually set No Show or Cancelled
      if (currentStatus !== "No Show" && currentStatus !== "Cancelled") {
        if (paymentVal === "Cash" || paymentVal === "Tikkie" || paymentVal === "Subscription" || paymentVal === "Free") {
          statusRange.setValue("Paid");
        } else if (paymentVal === "" && currentStatus === "Paid") {
          const dateVal = sheet.getRange(row, 2).getValue(); // col B = Date
          const isUpcoming = dateVal && new Date(dateVal) >= startOfDay_(new Date());
          statusRange.setValue(isUpcoming ? "Upcoming" : "Not Paid");
        }
      }

      // Price automation
      const priceRange = sheet.getRange(row, 5); // col E = Price
      const currentPrice = priceRange.getValue();
      const serviceName = String(sheet.getRange(row, 11).getValue()).toLowerCase(); // col K = Service

      if (paymentVal === "Subscription" || paymentVal === "Free") {
        priceRange.setValue(0);
      } else if (paymentVal === "" && currentPrice === 0) {
        priceRange.setValue(getStandardServicePrice_(serviceName));
      }
    }

    // B. Status changed (col G=7)
    if (range.getColumn() === 7) {
      const statusVal = range.getValue();

      // Clear Late checkbox ONLY for No Show / Cancelled
      if (statusVal === "No Show" || statusVal === "Cancelled") {
        sheet.getRange(row, 9).setValue(false); // col I = Late
      }

      // If setting to Free, ensure Price = 0 (but keep Late checkbox)
      if (statusVal.startsWith("Free")) {
        sheet.getRange(row, 5).setValue(0); // col E = Price
      }
    }

    // C. Recalculate all client stats immediately when any stat-affecting column changes.
    // Covers: E=Price, F=Payment, G=Status, H=Tips, I=Late, L=ClientID
    const col = range.getColumn();
    if (col === 5 || col === 6 || col === 7 || col === 8 || col === 9 || col === 12) {
      const minCtx = {
        appointmentsSheet: sheet,
        clientsSheet: sheet.getParent().getSheetByName(CLIENTS_SHEET)
      };
      updateClientStats_(minCtx);
      updateNoShowLateCounts_(minCtx);
      updateConsecutivePaidCounts_(minCtx);
    }
  }

  // 3. Client name auto-formatting (col B=2)
  if (sheetName === CLIENTS_SHEET && range.getColumn() === 2 && range.getRow() > HEADER_ROW) {
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
  const apptData = apptLastRow >= DATA_ROW ? ctx.appointmentsSheet.getRange(DATA_ROW, 2, apptLastRow - 2, 13).getValues() : [];
  updateConsecutivePaidCounts_(ctx, apptData);
  updateNoShowLateCounts_(ctx, apptData);
  updateClientStats_(ctx, apptData);
  sortAndHideAppointments_(ctx.appointmentsSheet);

  if (showNotification) {
    safeAlert_(`Sync Complete!\n\n+ ${counts.newCount} New\n~ ${counts.updatedCount} Updated\n- ${counts.cancelledCount} Cancelled`);
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
  const apptDataY = apptLastRowY >= DATA_ROW ? ctx.appointmentsSheet.getRange(DATA_ROW, 2, apptLastRowY - 2, 13).getValues() : [];
  updateConsecutivePaidCounts_(ctx, apptDataY);
  updateNoShowLateCounts_(ctx, apptDataY);
  updateClientStats_(ctx, apptDataY);
  sortAndHideAppointments_(ctx.appointmentsSheet);

  if (showNotification) {
    safeAlert_(`Sync Complete!\n\n+ ${counts.newCount} New\n~ ${counts.updatedCount} Updated\n- ${counts.cancelledCount} Cancelled`);
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
  const newEventIds = new Set();
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
        ctx.appointmentsSheet.getRange(existingRow, 7).setValue("Not Paid"); // col G = Status
      }

      // SEAMLESS SUBSCRIPTION CONVERSION
      if (isSubscriptionSale && oldPayment !== "Subscription") {
        ctx.appointmentsSheet.getRange(existingRow, 5).setValue(0);            // col E = Price
        ctx.appointmentsSheet.getRange(existingRow, 6).setValue("Subscription"); // col F = Payment
        ctx.appointmentsSheet.getRange(existingRow, 7).setValue("Paid");         // col G = Status
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

        // col D=4 Name formula; reading from col B so existingRow's L column = ClientID, N = CachedName
        const nameFormula = `=XLOOKUP(L${existingRow}; Clients!M:M; Clients!B:B; N${existingRow})`;
        ctx.appointmentsSheet.getRange(existingRow, 2, 1, 3).setValues([[dateCell, timeCell, nameFormula]]); // cols B-D
        ctx.appointmentsSheet.getRange(existingRow, 5).setValue(isSubscriptionSale ? 0 : rowPrice);          // col E = Price
        ctx.appointmentsSheet.getRange(existingRow, 10, 1, 5).setValues([[                                   // cols J-N
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
    const dateVal = entry.data[0]; // index 0 = col B = Date (range started at col 2)
    if (!dateVal) return;
    const rowDate = new Date(dateVal);
    if (rowDate < ctx.startDate || rowDate > ctx.endDate) return;

    const currentStatus = String(entry.data[5] || ""); // index 5 = col G = Status
    if (currentStatus === "Paid" || currentStatus === "No Show" || currentStatus === "Cancelled") return;

    ctx.appointmentsSheet.getRange(rowIdx, 7).setValue("Cancelled"); // col G = Status
    ctx.appointmentsSheet.getRange(rowIdx, 9).setValue(false);        // col I = Late
    cancelledCount++;
  });

  // Write all new rows
  if (newRowsABC.length > 0) {
    const startRow = ctx.appointmentsSheet.getLastRow() + 1;
    for (let i = 0; i < newRowsABC.length; i++) {
      const targetRow = startRow + i;
      // L = ClientID, N = CachedName in the new column layout
      newRowsABC[i][2] = `=XLOOKUP(L${targetRow}; Clients!M:M; Clients!B:B; N${targetRow})`;
    }
    ctx.appointmentsSheet.getRange(startRow, 2, newRowsABC.length, 3).setValues(newRowsABC);   // cols B-D
    ctx.appointmentsSheet.getRange(startRow, 5, newRowsDToM.length, 10).setValues(newRowsDToM); // cols E-N
  }

  return { newCount, updatedCount, cancelledCount, newEventIds };
}

/**
 * Flip any leftover "Upcoming" appointments whose date has passed to "Not Paid"
 */
function updateUpcomingToNotPaid_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_ROW) return;
  const today = startOfDay_(new Date());
  // Read cols B-G (Date through Status) — 6 columns starting at col 2
  const data = sheet.getRange(DATA_ROW, 2, lastRow - 2, 6).getValues();
  for (let i = 0; i < data.length; i++) {
    const dateVal = data[i][0]; // index 0 = col B = Date
    const status = data[i][5];  // index 5 = col G = Status
    if (!dateVal || status !== "Upcoming") continue;
    if (new Date(dateVal) < today) {
      sheet.getRange(i + DATA_ROW, 7).setValue("Not Paid"); // col G = Status
    }
  }
}

/**
 * Update Consecutive Paid counts in Clients sheet (col P=16)
 */
function updateConsecutivePaidCounts_(ctx, apptData) {
  const clientsSheet = ctx.clientsSheet;
  const apptSheet = ctx.appointmentsSheet;
  const lastRow = clientsSheet.getLastRow();
  if (lastRow < DATA_ROW) return;

  // Read 16 cols starting at col B (cols B-Q)
  const clientData = clientsSheet.getRange(DATA_ROW, 2, lastRow - 2, 16).getValues();
  if (!apptData) {
    const apptLastRow = apptSheet.getLastRow();
    apptData = apptLastRow >= DATA_ROW ? apptSheet.getRange(DATA_ROW, 2, apptLastRow - 2, 13).getValues() : [];
  }

  const n = clientData.length;
  const oVals = new Array(n);

  for (let i = 0; i < n; i++) {
    const clientId = clientData[i][11]; // index 11 = col M = ClientID
    if (!clientId) { oVals[i] = [0]; continue; }

    const clientAppts = [];
    for (const row of apptData) {
      // index 10 = col L = ClientID in appointments
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

  clientsSheet.getRange(DATA_ROW, 16, n, 1).setValues(oVals); // col P = ConsecutivePaid
}

/**
 * Update TotalVisits (J=10), TotalTips (K=11), TotalSpent (L=12), LastVisit (D=4), FirstVisit (N=14)
 * in Clients sheet.
 */
function updateClientStats_(ctx, apptData) {
  const clientsSheet = ctx.clientsSheet;
  const apptSheet = ctx.appointmentsSheet;
  const clientLastRow = clientsSheet.getLastRow();
  if (clientLastRow < DATA_ROW) return;

  // Read 12 cols starting at col B (cols B-M)
  const clientData = clientsSheet.getRange(DATA_ROW, 2, clientLastRow - 2, 12).getValues();
  if (!apptData) {
    const apptLastRow = apptSheet.getLastRow();
    apptData = apptLastRow >= DATA_ROW ? apptSheet.getRange(DATA_ROW, 2, apptLastRow - 2, 11).getValues() : [];
  }

  const lastVisits = [];   // col D = LastVisit
  const ijkValues = [];    // cols J, K, L (TotalVisits, TotalTips, TotalSpent)
  const firstVisits = [];  // col N = FirstVisit

  for (let i = 0; i < clientData.length; i++) {
    const clientId = clientData[i][11]; // index 11 = col M = ClientID
    if (!clientId) {
      lastVisits.push([""]);
      ijkValues.push([0, 0, 0]);
      firstVisits.push([""]);
      continue;
    }

    let totalVisits = 0, totalTips = 0, totalSpent = 0;
    let firstVisit = null, lastVisit = null;

    for (const row of apptData) {
      if (String(row[10]) !== String(clientId)) continue; // index 10 = col L = ClientID

      const payment = String(row[4] || "");  // index 4 = col F = Payment
      const status  = String(row[5] || "");  // index 5 = col G = Status
      const isPaidVisit = payment === "Cash" || payment === "Tikkie" ||
                          payment === "Subscription" || payment === "Free" ||
                          status.startsWith("Free");
      if (!isPaidVisit) continue;

      const dateVal = row[0]; // index 0 = col B = Date
      if (dateVal) {
        const d = new Date(dateVal);
        if (!firstVisit || d < firstVisit) firstVisit = d;
        if (!lastVisit  || d > lastVisit)  lastVisit  = d;
      }

      totalVisits++;
      totalTips += Number(row[6]) || 0; // index 6 = col H = Tips
      if (payment === "Cash" || payment === "Tikkie" || payment === "Subscription") {
        totalSpent += Number(row[3]) || 0; // index 3 = col E = Price
      }
    }

    lastVisits.push([lastVisit  || ""]);
    ijkValues.push([totalVisits, totalTips, totalSpent]);
    firstVisits.push([firstVisit || ""]);
  }

  const n = clientData.length;
  clientsSheet.getRange(DATA_ROW, 4,  n, 1).setValues(lastVisits);  // col D = LastVisit
  clientsSheet.getRange(DATA_ROW, 10, n, 3).setValues(ijkValues);   // cols J, K, L
  clientsSheet.getRange(DATA_ROW, 14, n, 1).setValues(firstVisits); // col N = FirstVisit
}

/**
 * Update NoShow (col G=7) and Late (col H=8) counts in Clients sheet.
 */
function updateNoShowLateCounts_(ctx, apptData) {
  const clientsSheet = ctx.clientsSheet;
  const apptSheet = ctx.appointmentsSheet;
  const clientLastRow = clientsSheet.getLastRow();
  if (clientLastRow < DATA_ROW) return;

  if (!apptData) {
    const apptLastRow = apptSheet.getLastRow();
    if (apptLastRow < DATA_ROW) return;
    apptData = apptSheet.getRange(DATA_ROW, 2, apptLastRow - 2, 11).getValues();
  }

  const cutoff = new Date();
  cutoff.setFullYear(cutoff.getFullYear() - 1);

  // Read 12 cols from col B (B-M)
  const clientData = clientsSheet.getRange(DATA_ROW, 2, clientLastRow - 2, 12).getValues();
  const n = clientData.length;
  const fgVals = new Array(n);

  for (let i = 0; i < n; i++) {
    const clientId = clientData[i][11]; // index 11 = col M = ClientID
    if (!clientId) { fgVals[i] = [0, 0]; continue; }

    let noShows = 0, lates = 0;
    for (const row of apptData) {
      if (String(row[10]) !== String(clientId)) continue; // index 10 = col L = ClientID
      const dateVal = row[0];
      if (!dateVal || new Date(dateVal) < cutoff) continue;
      if (row[5] === "No Show") noShows++; // index 5 = col G = Status
      if (row[7] === true) lates++;        // index 7 = col I = Late
    }
    fgVals[i] = [noShows, lates];
  }

  clientsSheet.getRange(DATA_ROW, 7, n, 2).setValues(fgVals); // cols G-H = NoShow, Late
}

function sortAndHideAppointments_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_ROW) return;

  // Sort data rows (starting at DATA_ROW, from col B). Column numbers are relative to the range.
  // Column 1 of range = col B = Date, Column 2 = col C = Time.
  sheet.getRange(DATA_ROW, 2, lastRow - 2, sheet.getLastColumn() - 1).sort([
    { column: 1, ascending: false },
    { column: 2, ascending: false }
  ]);

  // Read just the Date column (col B) to determine which rows to hide
  const dates = sheet.getRange(DATA_ROW, 2, lastRow - 2, 1).getValues();
  const today = new Date();
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

  sheet.showRows(DATA_ROW, lastRow - 2);
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
  if (lastRow >= DATA_ROW) {
    // Check last ~20 rows to avoid duplicate subscription entries
    const checkStart = Math.max(DATA_ROW, lastRow - 20);
    const checkCount = lastRow - checkStart + 1;
    const checkRange = subsSheet.getRange(checkStart, 2, checkCount, 4).getValues();
    const startYMD = ymd_(startDate instanceof Date ? startDate : new Date(startDate), ctx.calTz);
    for (const r of checkRange) {
      if (nameCase_(r[0]) === name && r[3] && ymd_(new Date(r[3]), ctx.calTz) === startYMD) return;
    }
  }

  const r = lastRow + 1;
  // ClientID is in col J of Subscriptions (index 8 of the 9-col data, col B=Name through J=ClientID)
  // Clients!M:M = ClientID, Clients!B:B = Name
  const nameFormula = `=XLOOKUP(J${r}; Clients!M:M; Clients!B:B; "")`;
  // Subscriptions layout (starting col B): B=Name C=Price D=? E=StartDate F=Credits G=Status H=Expiry I=? J=ClientID
  // E${r}=StartDate, H${r}=Expiry; Appointments cols: $L:$L=ClientID, $F:$F=Payment, $G:$G=Status, $B:$B=Date
  const creditsFormula = `=MAX(0; 4 - COUNTIFS(Appointments!$L:$L; J${r}; Appointments!$F:$F; "Subscription"; Appointments!$G:$G; "Paid"; Appointments!$B:$B; ">="&E${r}; Appointments!$B:$B; "<="&(H${r} + 21)))`;

  // Write to cols B-J (9 columns, starting at col 2) — col A is the spacer
  subsSheet.getRange(r, 2, 1, 9).setValues([[nameFormula, monthlyPrice, "", startDate, creditsFormula, "Active", addDays_(startDate, 31), "", clientId]]);

  const newLastRow = subsSheet.getLastRow();
  if (newLastRow >= DATA_ROW) {
    // Sort by StartDate (col E = column 4 within range starting at col B)
    subsSheet.getRange(DATA_ROW, 2, newLastRow - 2, 9).sort({ column: 4, ascending: false });
  }

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

  if (lastRow < DATA_ROW) { safeAlert_("No appointment data to migrate."); return; }

  // Ensure col N header exists (CachedName, col 14)
  if (!sheet.getRange(HEADER_ROW, 14).getValue()) {
    sheet.getRange(HEADER_ROW, 14).setValue("Cached Name");
  }

  // Read 13 cols from col B (B-N)
  const data = sheet.getRange(DATA_ROW, 2, lastRow - 2, 13).getValues();
  const cachedNames = [];
  const updatedFormulas = [];

  for (let i = 0; i < data.length; i++) {
    const row = i + DATA_ROW;
    const currentName = String(data[i][2] || "");   // index 2 = col D = Name
    const existingCache = String(data[i][12] || ""); // index 12 = col N = CachedName

    cachedNames.push([existingCache || currentName]);
    // L = ClientID, N = CachedName in new layout
    updatedFormulas.push([`=XLOOKUP(L${row}; Clients!M:M; Clients!B:B; N${row})`]);
  }

  sheet.getRange(DATA_ROW, 14, cachedNames.length, 1).setValues(cachedNames);   // col N = CachedName
  sheet.getRange(DATA_ROW, 4, updatedFormulas.length, 1).setFormulas(updatedFormulas); // col D = Name formula

  safeAlert_(`✅ Migration complete!\n${cachedNames.length} rows updated.\n\nYou can now hide column N (right-click column N header → Hide column).`);
}

/***************
 * HELPERS
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
  // Read 12 cols from col B (B-M): index 0=Name, index 11=ClientID
  const data = sheet.getRange(DATA_ROW, 2, lastRow - 2, 12).getValues();
  for (let i = 0; i < data.length; i++) {
    const name = nameCase_(data[i][0]);   // index 0 = col B = Name
    const id = data[i][11];               // index 11 = col M = ClientID
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
  if (sheet.getMaxColumns() < 18) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), 18 - sheet.getMaxColumns());
  }
  sheet.getRange(newRow, 2).setValue(formatted);   // col B = Name
  sheet.getRange(newRow, 13).setValue(idx.maxId);  // col M = ClientID
  idx.map.set(formatted, idx.maxId);
  return idx.maxId;
}

function loadActiveSubscriptionsIndex_(sheet) {
  const today = startOfDay_(new Date());
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_ROW) return { byId: new Map(), byName: new Map() };
  // Subscriptions layout from col B: B=Name(0) C=Price(1) D=?(2) E=StartDate(3) F=Credits(4) G=Status(5) H=Expiry(6) I=?(7) J=ClientID(8)
  const data = sheet.getRange(DATA_ROW, 2, lastRow - 2, 9).getValues();
  const byId = new Map(), byName = new Map();
  for (let i = 0; i < data.length; i++) {
    const name     = data[i][0]; // col B = Name
    const credits  = data[i][4]; // col F = Credits
    const status   = data[i][5]; // col G = Status
    const expiryRaw = data[i][6]; // col H = Expiry
    const clientId = data[i][8]; // col J = ClientID
    const expiry = expiryRaw ? new Date(expiryRaw) : null;
    if (status !== "Active" || !expiry || expiry < today) continue;
    if (credits !== "" && Number(credits) <= 0) continue;
    const entry = { start: data[i][3] ? new Date(data[i][3]) : new Date(0) }; // col E = StartDate
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
  // Read 13 cols from col B (B-N): index 11=EventID (col M), index 10=ClientID (col L)
  const data = sheet.getRange(DATA_ROW, 2, lastRow - 2, 13).getValues();
  for (let i = 0; i < data.length; i++) {
    const eventId = data[i][11]; // index 11 = col M = EventID
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
 */
function setupIncrementalSync() {
  removeIncrementalSync();
  PropertiesService.getScriptProperties().deleteProperty("CALENDAR_SYNC_TOKEN");
  ScriptApp.newTrigger("syncCalendarIncremental_")
    .timeBased().everyMinutes(5).create();
  syncCalendarIncremental_();
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
 * ONE-CLICK SETUP — installs onEdit trigger + 5-min incremental sync.
 */
function setupTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => {
    const fn = t.getHandlerFunction();
    if (fn === "processSheetChanges" || fn === "syncCalendarIncremental_") {
      ScriptApp.deleteTrigger(t);
    }
  });
  PropertiesService.getScriptProperties().deleteProperty("CALENDAR_SYNC_TOKEN");

  ScriptApp.newTrigger("processSheetChanges").forSpreadsheet(SpreadsheetApp.getActive()).onEdit().create();
  ScriptApp.newTrigger("syncCalendarIncremental_").timeBased().everyMinutes(5).create();

  syncCalendarIncremental_();
  safeAlert_("✅ Triggers installed!\n\n• onEdit → processSheetChanges\n• Every 5 min → syncCalendarIncremental_\n\nIf you see a Calendar error go to:\nServices (+) → Google Calendar API → Add");
}

/**
 * Installs an installable onOpen trigger so the sheet syncs automatically when opened.
 */
function setupOnOpenSync() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "onOpenSync_") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("onOpenSync_").forSpreadsheet(SpreadsheetApp.getActive()).onOpen().create();
  safeAlert_("✅ Sync on sheet open enabled.");
}

function onOpenSync_() {
  try { syncCalendarIncremental_(); } catch (e) { Logger.log("onOpen sync error: " + e.message); }
}

/**
 * Checks that everything is configured correctly.
 */
function validateSetup() {
  const lines = [];

  try {
    getCalendarOrThrow_();
    lines.push("✅ Calendar \"" + CALENDAR_NAME + "\" found");
  } catch (e) {
    lines.push("❌ Calendar: " + e.message);
  }

  const ss = SpreadsheetApp.getActive();
  [APPOINTMENTS_SHEET, CLIENTS_SHEET, SERVICES_SHEET, SUBSCRIPTIONS_SHEET].forEach(name => {
    lines.push(ss.getSheetByName(name) ? "✅ Sheet \"" + name + "\" found" : "❌ Sheet \"" + name + "\" MISSING");
  });

  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty("TELEGRAM_BOT_TOKEN");
  const chatId = props.getProperty("TELEGRAM_CHAT_ID");
  lines.push(token ? "✅ TELEGRAM_BOT_TOKEN set" : "❌ TELEGRAM_BOT_TOKEN not set");
  lines.push(chatId ? "✅ TELEGRAM_CHAT_ID set" : "❌ TELEGRAM_CHAT_ID not set");

  const triggers = ScriptApp.getProjectTriggers().map(t => t.getHandlerFunction());
  lines.push(triggers.includes("processSheetChanges") ? "✅ onEdit trigger active" : "⚠️ onEdit trigger missing — run 🛠️ Setup All Triggers");
  lines.push(triggers.includes("syncCalendarIncremental_") ? "✅ 5-min incremental sync active" : "⚠️ 5-min sync missing — run ⚡ Setup Incremental Sync");
  lines.push(triggers.includes("sendDailyNotification") ? "✅ Daily notification trigger active" : "⚠️ Daily notification missing — run setupNotificationTrigger()");
  lines.push(triggers.includes("onOpenSync_") ? "✅ onOpen sync active" : "ℹ️ onOpen sync not set (optional — run 📲 Setup Sync on Sheet Open)");

  const syncToken = props.getProperty("CALENDAR_SYNC_TOKEN");
  lines.push(syncToken ? "✅ Calendar sync token present" : "ℹ️ No sync token yet (runs full sync on first trigger fire)");

  safeAlert_("Setup validation:\n\n" + lines.join("\n"));
  Logger.log(lines.join("\n"));
}

function syncCalendarIncremental_() {
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
    const ss           = SpreadsheetApp.getActive();
    const apptSheet    = getSheetOrThrow_(ss, APPOINTMENTS_SHEET);
    const clientsSheet = getSheetOrThrow_(ss, CLIENTS_SHEET);
    const servicesSheet = getSheetOrThrow_(ss, SERVICES_SHEET);
    const subsSheet    = getSheetOrThrow_(ss, SUBSCRIPTIONS_SHEET);
    const today        = startOfDay_(new Date());

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
        const s = String(entry.data[5] || ""); // index 5 = col G = Status
        if (s !== "Paid" && s !== "No Show" && s !== "Cancelled") {
          apptSheet.getRange(rowIdx, 7).setValue("Cancelled"); // col G = Status
          apptSheet.getRange(rowIdx, 9).setValue(false);        // col I = Late
          hasChanges = true;
        }
        continue;
      }

      const parsed = parseEventTitle_(item.summary || "");
      if (!parsed) continue;

      const clientId   = getOrCreateClientId_(clientsSheet, ctx.clientsIndex, parsed.clientName);
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
          const formula = `=XLOOKUP(L${existingRow}; Clients!M:M; Clients!B:B; N${existingRow})`;
          apptSheet.getRange(existingRow, 2, 1, 3).setValues([[dateCell, timeCell, formula]]); // cols B-D
          apptSheet.getRange(existingRow, 10, 1, 5).setValues([[                               // cols J-N
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
        newRowsABC[i][2] = `=XLOOKUP(L${r}; Clients!M:M; Clients!B:B; N${r})`;
      }
      apptSheet.getRange(startRow, 2, newRowsABC.length, 3).setValues(newRowsABC);   // cols B-D
      apptSheet.getRange(startRow, 5, newRowsDToM.length, 10).setValues(newRowsDToM); // cols E-N
    }

    updateUpcomingToNotPaid_(apptSheet);
    const iLastRow = apptSheet.getLastRow();
    const iApptData = iLastRow >= DATA_ROW ? apptSheet.getRange(DATA_ROW, 2, iLastRow - 2, 13).getValues() : [];
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
 * Safe: only deletes rows where Payment is empty AND status is Not Paid/Upcoming
 * AND another row exists for the same client at the same date+time.
 */
function cleanupDuplicates() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(APPOINTMENTS_SHEET);
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_ROW) { safeAlert_("No data found."); return; }

  // Read 13 cols from col B (B-N)
  const data = sheet.getRange(DATA_ROW, 2, lastRow - 2, 13).getValues();

  const groups = new Map();
  for (let i = 0; i < data.length; i++) {
    const row     = data[i];
    const dateStr = row[0] ? new Date(row[0]).toDateString() : ""; // index 0 = col B = Date
    const timeStr = row[1] ? String(row[1]) : "";                  // index 1 = col C = Time
    const name    = String(row[2] || row[12] || "").trim().toLowerCase(); // col D=Name, col N=CachedName
    const payment = String(row[4] || "").trim(); // index 4 = col F = Payment
    const status  = String(row[5] || "").trim(); // index 5 = col G = Status

    if (!dateStr || !name) continue;

    const key = `${dateStr}|${timeStr}|${name}`;
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push({ sheetRow: i + DATA_ROW, payment, status });
  }

  const toDelete = [];
  for (const rows of groups.values()) {
    if (rows.length <= 1) continue;

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

  toDelete.sort((a, b) => b - a);
  for (const rowIdx of toDelete) {
    sheet.deleteRow(rowIdx);
  }

  safeAlert_(`✅ Removed ${toDelete.length} duplicate rows.\n\nNow run: ✂️ Barber Tools → Setup Incremental Sync`);
}

function doGet(e) {
  try {
    syncCalendarIncremental_();
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
 * Applies the visual theme to all sheets.
 * Run from the Apps Script editor dropdown. Safe to re-run at any time.
 */
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

/**
 * Sets the base theme on a sheet. Tables start at B2: row 1 and col A are spacers.
 * Returns the alternating-row rule to be appended last (lowest priority) by the caller.
 */
function applyBaseTheme_(sheet, COLORS) {
  if (!sheet) return [];

  Logger.log('Formatting: ' + sheet.getName());

  const maxRows = sheet.getMaxRows();
  const maxCols = Math.min(Math.max(sheet.getLastColumn(), 1), sheet.getMaxColumns());
  const styledRows = Math.min(Math.max(sheet.getLastRow(), 1) + 10, maxRows);

  sheet.clearConditionalFormatRules();
  sheet.showColumns(1, maxCols);

  // Base background for all cells (including spacer row/col)
  sheet.getRange(1, 1, styledRows, maxCols)
    .setBackground(COLORS.bg)
    .setFontColor(COLORS.text)
    .setFontWeight('normal');

  // Header row (row 2): styled tan background, bold text
  sheet.getRange(HEADER_ROW, 1, 1, maxCols)
    .setBackground(COLORS.headerBg)
    .setFontColor(COLORS.text)
    .setFontWeight('bold');

  // Freeze both the spacer row and the header row so they stay visible while scrolling
  sheet.setFrozenRows(HEADER_ROW);

  // Row 1 = spacer: keep slim. Row 2 = header: tall so it reads as a clear label at 25% zoom.
  sheet.setRowHeight(1, 20);
  sheet.setRowHeight(HEADER_ROW, 320);
  if (maxRows > HEADER_ROW) {
    sheet.setRowHeightsForced(DATA_ROW, maxRows - DATA_ROW + 1, 280);
  }

  // Alternating row rule — applied only to data rows, lowest priority so status colors win
  const altRule = SpreadsheetApp.newConditionalFormatRule()
    .withCriteria(SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA, ['=MOD(ROW(),2)=0'])
    .setBackground(COLORS.surface)
    .setRanges([sheet.getRange(DATA_ROW, 1, Math.max(maxRows - DATA_ROW + 1, 1), maxCols)])
    .build();

  return [altRule];
}

/**
 * Formats the Appointments sheet.
 * Cols (after spacer A): B=Date C=Time D=Name E=Price F=Payment G=Status H=Tips I=Late J=Notes K=Service L=ClientID M=EventID N=CachedName
 */
function formatAppointments_(sheet, COLORS) {
  if (!sheet) return;

  const baseRules = applyBaseTheme_(sheet, COLORS);

  // Col A = spacer (20). Then B-K visible, L-N hidden.
  // B=Date(140) C=Time(110) D=Name(290) E=Price(110) F=Payment(185) G=Status(155)
  // H=Tips(110) I=Late(90) J=Notes(240) K=Service(210) → visible total ~1640px
  const widths = [20, 140, 110, 290, 110, 185, 155, 110, 90, 240, 210, 80, 80, 80];
  widths.forEach((w, i) => { if (w > 0) sheet.setColumnWidth(i + 1, w); });

  // Hide internal columns: L=ClientID(12) M=EventID(13) N=CachedName(14)
  sheet.hideColumns(12, 3);

  const maxRows = sheet.getMaxRows();
  const dataRows = Math.max(maxRows - DATA_ROW + 1, 1);
  // Data range starts at B3 (col 2, row DATA_ROW), 13 columns wide (B-N)
  const dataRange = sheet.getRange(DATA_ROW, 2, dataRows, 13);

  // Typography — CLIP stops text wrapping; sizes target ~10-12pt visible at 25% zoom
  dataRange
    .setFontSize(26)
    .setVerticalAlignment('middle')
    .setFontWeight('normal')
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

  sheet.getRange(HEADER_ROW, 2, 1, 10).setFontSize(18); // header labels B-K

  // Hero columns: Time (C=3) and Name (D=4) — bold for visual hierarchy
  sheet.getRange(DATA_ROW, 3, dataRows, 1).setFontSize(30).setFontWeight('bold'); // Time
  sheet.getRange(DATA_ROW, 4, dataRows, 1).setFontSize(30).setFontWeight('bold').setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP); // Name

  // Date (B=2): slightly smaller so the full date fits without clipping
  sheet.getRange(DATA_ROW, 2, dataRows, 1).setFontSize(20);

  // Secondary info: muted colour
  sheet.getRange(DATA_ROW, 10, dataRows, 1).setFontSize(22).setFontColor(COLORS.textMuted); // Notes J=10
  sheet.getRange(DATA_ROW, 11, dataRows, 1).setFontSize(22).setFontColor(COLORS.textMuted); // Service K=11

  // Horizontal alignment
  const centred = [2, 3, 5, 6, 7, 8, 9]; // Date, Time, Price, Payment, Status, Tips, Late
  centred.forEach(col => sheet.getRange(DATA_ROW, col, dataRows, 1).setHorizontalAlignment('center'));

  // Conditional format rules (index 0 = highest priority)
  // Formula row reference uses DATA_ROW (3) since that's the first row of the applied range.
  // Column letters reflect the new layout: I=Late, G=Status, F=Payment.
  const ruleDefs = [
    { formula: `=$I${DATA_ROW}=TRUE`,              bg: COLORS.tint.orange, fg: COLORS.accent.orange }, // Late
    { formula: `=$G${DATA_ROW}="No Show"`,         bg: COLORS.tint.red,    fg: COLORS.accent.red    },
    { formula: `=$G${DATA_ROW}="Cancelled"`,       bg: COLORS.bg,          fg: COLORS.textMuted     },
    { formula: `=$F${DATA_ROW}="Subscription"`,    bg: COLORS.tint.purple, fg: COLORS.accent.purple },
    { formula: `=$G${DATA_ROW}="Paid"`,            bg: COLORS.tint.green,  fg: COLORS.accent.green  },
    { formula: `=$G${DATA_ROW}="Upcoming"`,        bg: COLORS.tint.blue,   fg: COLORS.accent.blue   },
    { formula: `=$G${DATA_ROW}="Not Paid"`,        bg: COLORS.tint.yellow, fg: COLORS.accent.yellow },
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
  sheet.setTabColor(COLORS.accent.green);
}

/**
 * Formats the Clients sheet.
 * Cols (after spacer A): B=Name C=FavService D=LastVisit E=SocialMedia F=Notes G=NoShow H=Late
 *                        I=Referral J=TotalVisits K=TotalTips L=TotalSpent M=ClientID N=FirstVisit
 *                        O=DoNotCut P=ConsecutivePaid Q=VIP
 */
function formatClients_(sheet, COLORS) {
  if (!sheet) return;

  const baseRules = applyBaseTheme_(sheet, COLORS);

  // Col A = spacer (20). Visible after hiding: B C D F G H O P Q = 9 cols ~1620px
  // A=spacer(20) B=Name(280) C=FavService(240) D=LastVisit(170) E=hidden(100) F=Notes(310)
  // G=NoShow(140) H=Late(120) I-N=hidden(100ea) O=DoNotCut(130) P=Consec(130) Q=VIP(100)
  const widths = [20, 280, 240, 170, 100, 310, 140, 120, 100, 100, 100, 100, 100, 100, 130, 130, 100];
  widths.forEach((w, i) => { if (w > 0) sheet.setColumnWidth(i + 1, w); });

  // Hide: E=SocialMedia(5), I=Referral(9), J=TotalVisits(10), K=TotalTips(11),
  //       L=TotalSpent(12), M=ClientID(13), N=FirstVisit(14)
  sheet.hideColumns(5, 1);   // E SocialMedia
  sheet.hideColumns(9, 6);   // I–N (Referral, TotalVisits, TotalTips, TotalSpent, ClientID, FirstVisit)

  const maxRows = sheet.getMaxRows();
  const dataRows = Math.max(maxRows - DATA_ROW + 1, 1);
  // Data range: B3 (col 2, row DATA_ROW), 16 columns (B-Q)
  const dataRange = sheet.getRange(DATA_ROW, 2, dataRows, 16);

  dataRange
    .setFontSize(26)
    .setVerticalAlignment('middle')
    .setFontWeight('normal')
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  sheet.getRange(HEADER_ROW, 2, 1, 16).setFontSize(18);

  // Name (B=2) is the hero column
  sheet.getRange(DATA_ROW, 2, dataRows, 1).setFontSize(30).setFontWeight('bold').setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

  // Centre-align: LastVisit(D=4), NoShow(G=7), Late(H=8), DoNotCut(O=15), Consec(P=16), VIP(Q=17)
  [4, 7, 8, 15, 16, 17].forEach(col =>
    sheet.getRange(DATA_ROW, col, dataRows, 1).setHorizontalAlignment('center')
  );

  sheet.getRange(DATA_ROW, 3, dataRows, 1).setFontSize(22); // FavService C=3
  sheet.getRange(DATA_ROW, 6, dataRows, 1).setFontSize(22).setFontColor(COLORS.textMuted); // Notes F=6

  const ruleDefs = [
    { formula: `=$O${DATA_ROW}=TRUE`,              bg: COLORS.tint.red,    fg: COLORS.accent.red    }, // DoNotCut O
    { formula: `=$G${DATA_ROW}+$H${DATA_ROW}>=3`, bg: COLORS.tint.orange, fg: COLORS.accent.orange }, // NoShow+Late ≥ 3
    { formula: `=$Q${DATA_ROW}=TRUE`,              bg: COLORS.tint.yellow, fg: COLORS.accent.yellow }, // VIP Q
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
  sheet.setTabColor(COLORS.accent.purple);
}

/** Formats the Services sheet. */
function formatServices_(sheet, COLORS) {
  if (!sheet) return;
  const baseRules = applyBaseTheme_(sheet, COLORS);
  sheet.setColumnWidth(1, 20);  // spacer col A
  sheet.setColumnWidth(2, 150); // col B = ServiceName
  sheet.setColumnWidth(3, 90);  // col C = Price
  sheet.setConditionalFormatRules(baseRules);
  sheet.setTabColor(COLORS.textMuted);
}

/** Formats the Subscriptions sheet. */
function formatSubscriptions_(sheet, COLORS) {
  if (!sheet) return;
  const baseRules = applyBaseTheme_(sheet, COLORS);
  sheet.setConditionalFormatRules(baseRules);
  sheet.setTabColor(COLORS.accent.orange);
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

  sheet.getRange(1, 2, styledRows, 1).setFontColor(COLORS.textMuted);
  [3, 8, 9, 10].forEach(col => {
    if (col <= styledCols) sheet.getRange(1, col, styledRows, 1).setFontWeight('bold').setFontColor(COLORS.text);
  });
  if (styledCols >= 10) sheet.getRange('J2').setFontColor(COLORS.accent.blue).setFontWeight('bold');

  if (styledCols >= 3) {
    sheet.getRange(3, 1, 1, styledCols)
      .setBackground(COLORS.tint.blue)
      .setFontWeight('bold');
  }

  sheet.setConditionalFormatRules(baseRules);
  sheet.setTabColor(COLORS.accent.blue);
}
