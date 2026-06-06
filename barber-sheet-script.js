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
// L=12 CachedName  M=13 ClientID  N=14 EventID  O=15 spacer

// Clients columns (1-based):
// B=2 Name  C=3 FavService  D=4 LastVisit  E=5 SocialMedia  F=6 Notes
// G=7 NoShow(12m)  H=8 Late(12m)  I=9 Referral  J=10 TotalVisits
// K=11 TotalSpent  L=12 TotalTips  M=13 FirstVisit  N=14 ConsecutivePaid
// O=15 VIP  P=16 DoNotCut  Q=17 ClientID  R=18 spacer

// Subscriptions columns (1-based):
// B=2 Name  C=3 Price  D=4 Type  E=5 Expiry  F=6 Credits  G=7 Status
// H=8 Notes  I=9 StartDate  J=10 ClientID  K=11 spacer

// Auto-VIP thresholds (used in updateClientStats_)
const VIP_MIN_VISITS = 15;  // total paid visits to earn VIP
const VIP_MIN_SPENT  = 400; // total € spent to earn VIP

/***************
 * TRIGGERS
 ***************/

/**
 * INSTALLABLE TRIGGER — linked to 'On edit' in Script Settings
 */
function processSheetChanges(e) {
  try {
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

    // 2. Dashboard period selector (C4) — sets start date (C5) and end date (C6)
    if (sheetName === "Dashboard" && range.getRow() === 4 && range.getColumn() === 3) {
      const period = range.getValue();
      const now = new Date();
      let startDate = null;
      let endDate   = null;

      if (period === "Today") {
        startDate = startOfDay_(now);
        endDate   = endOfDay_(now);

      } else if (period === "This Week") {
        const day = now.getDay();                        // 0=Sun … 6=Sat
        const diffToMon = (day === 0 ? -6 : 1 - day);  // days back to Monday
        startDate = startOfDay_(addDays_(now, diffToMon));
        endDate   = endOfDay_(addDays_(startDate, 6));

      } else if (period === "This Month") {
        startDate = new Date(now.getFullYear(), now.getMonth(), 1);
        endDate   = endOfDay_(new Date(now.getFullYear(), now.getMonth() + 1, 0));

      } else if (period === "Last Month") {
        startDate = new Date(now.getFullYear(), now.getMonth() - 1, 1);
        endDate   = endOfDay_(new Date(now.getFullYear(), now.getMonth(), 0));

      } else if (period === "Last 3 Months") {
        startDate = new Date(now.getFullYear(), now.getMonth() - 3, 1);
        endDate   = endOfDay_(now);

      } else if (period === "YTD") {
        startDate = new Date(now.getFullYear(), 0, 1);  // Jan 1 this year
        endDate   = endOfDay_(now);

      } else if (period === "All Time") {
        startDate = new Date(2020, 0, 1);
        endDate   = endOfDay_(now);

      } else if (period === "Custom") {
        sheet.getRange(5, 3).clearContent();
        sheet.getRange(6, 3).clearContent();
        return;
      }

      if (startDate && endDate) {
        sheet.getRange(5, 3).setValue(startDate);
        sheet.getRange(6, 3).setValue(endDate);
      }
      return;
    }

    // 3. Appointments sheet logic
    if (sheetName === APPOINTMENTS_SHEET && range.getRow() > HEADER_ROW) {
      const numRows = range.getNumRows();

      // A. Payment type changed (col F=6) — handles multi-row paste
      if (range.getColumn() === 6) {
        const values = range.getValues(); // 2D array, one row per pasted cell
        for (let i = 0; i < numRows; i++) {
          const rowNum = range.getRow() + i;
          const paymentVal = values[i][0];
          const statusRange = sheet.getRange(rowNum, 7); // col G = Status
          const currentStatus = statusRange.getValue();

          // Never override a manually set No Show or Cancelled
          if (currentStatus !== "No Show" && currentStatus !== "Cancelled") {
            if (paymentVal === "Cash" || paymentVal === "Tikkie" || paymentVal === "Subscription" || paymentVal === "Free") {
              statusRange.setValue("Paid");
            } else if (paymentVal === "" && currentStatus === "Paid") {
              const dateVal = sheet.getRange(rowNum, 2).getValue(); // col B = Date
              const isUpcoming = dateVal && new Date(dateVal) >= startOfDay_(new Date());
              statusRange.setValue(isUpcoming ? "Upcoming" : "Not Paid");
            }
          }

          // Price automation
          const priceRange = sheet.getRange(rowNum, 5); // col E = Price
          const currentPrice = priceRange.getValue();
          const serviceName = String(sheet.getRange(rowNum, 11).getValue()).toLowerCase(); // col K = Service

          if (paymentVal === "Subscription" || paymentVal === "Free") {
            priceRange.setValue(0);
          } else if (paymentVal === "" && currentPrice === 0) {
            priceRange.setValue(getStandardServicePrice_(serviceName));
          }
        }
      }

      // B. Status changed (col G=7) — handles multi-row paste
      if (range.getColumn() === 7) {
        const values = range.getValues();
        for (let i = 0; i < numRows; i++) {
          const rowNum = range.getRow() + i;
          const statusVal = values[i][0];

          // Clear Late checkbox ONLY for No Show / Cancelled
          if (statusVal === "No Show" || statusVal === "Cancelled") {
            sheet.getRange(rowNum, 9).setValue(false); // col I = Late
          }

          // Unify Free-* → canonical state: Payment=Free, Status=Paid, Price=0
          if (String(statusVal).startsWith("Free")) {
            sheet.getRange(rowNum, 5).setValue(0);        // col E = Price
            sheet.getRange(rowNum, 6).setValue("Free");   // col F = Payment
            sheet.getRange(rowNum, 7).setValue("Paid");   // col G = Status
          }
        }
      }

      // C. Recalculate all client stats immediately when any stat-affecting column changes.
      // Covers: E=Price, F=Payment, G=Status, H=Tips, I=Late, M=ClientID
      const col = range.getColumn();
      if (col === 5 || col === 6 || col === 7 || col === 8 || col === 9 || col === 13) {
        const minCtx = {
          appointmentsSheet: sheet,
          clientsSheet: sheet.getParent().getSheetByName(CLIENTS_SHEET)
        };
        // Read appointment data once and share across all three update functions
        const apptLastRow = sheet.getLastRow();
        const apptData = apptLastRow >= DATA_ROW
          ? sheet.getRange(DATA_ROW, 2, apptLastRow - 2, 13).getValues()
          : [];
        updateClientStats_(minCtx, apptData);
        updateNoShowLateCounts_(minCtx, apptData);
        updateConsecutivePaidCounts_(minCtx, apptData);
      }
    }

    // 4. Client name auto-formatting (col B=2)
    if (sheetName === CLIENTS_SHEET && range.getColumn() === 2 && range.getRow() > HEADER_ROW) {
      const v = range.getValue();
      const fixed = nameCase_(v);
      if (fixed !== v) range.setValue(fixed);
    }
  } catch (err) {
    Logger.log("processSheetChanges error: " + err.message + "\n" + err.stack);
    sendTelegramError_("onEdit error: " + err.message);
  }
}

/**
 * Web app endpoint — called by iOS Shortcut when Calendar app closes.
 */
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

/***************
 * SETUP & TRIGGERS
 ***************/

/**
 * ONE-CLICK SETUP — installs onEdit trigger + 5-min incremental sync + midnight sweep.
 */
function setupTriggers() {
  try {
    Logger.log("setupTriggers: removing existing triggers");
    ScriptApp.getProjectTriggers().forEach(t => {
      const fn = t.getHandlerFunction();
      if (fn === "processSheetChanges" || fn === "syncCalendarIncremental_" || fn === "runMidnightSweep_") {
        ScriptApp.deleteTrigger(t);
      }
    });
    PropertiesService.getScriptProperties().deleteProperty("CALENDAR_SYNC_TOKEN");

    ScriptApp.newTrigger("processSheetChanges").forSpreadsheet(SpreadsheetApp.getActive()).onEdit().create();
    ScriptApp.newTrigger("syncCalendarIncremental_").timeBased().everyMinutes(5).create();
    ScriptApp.newTrigger("runMidnightSweep_").timeBased().everyDays(1).atHour(0).create();
    Logger.log("setupTriggers: triggers created, running initial sync");

    syncCalendarIncremental_();
    Logger.log("setupTriggers complete");
    notify_("✅ Triggers installed!\n\n• onEdit → processSheetChanges\n• Every 5 min → syncCalendarIncremental_\n• Daily at midnight → runMidnightSweep_\n\nIf you see a Calendar error go to:\nServices (+) → Google Calendar API → Add", 8);
  } catch (e) {
    Logger.log("setupTriggers error: " + e.message + "\n" + e.stack);
    sendTelegramError_("Setup triggers failed: " + e.message);
    throw e;
  }
}

function setupIncrementalSync() {
  try {
    Logger.log("setupIncrementalSync: removing existing triggers");
    removeIncrementalSync();
    PropertiesService.getScriptProperties().deleteProperty("CALENDAR_SYNC_TOKEN");
    ScriptApp.newTrigger("syncCalendarIncremental_").timeBased().everyMinutes(5).create();
    Logger.log("setupIncrementalSync: trigger created, running initial sync");
    syncCalendarIncremental_();
    Logger.log("setupIncrementalSync complete");
    notify_("✅ Incremental sync running every 5 minutes.\n\nIf you see a Calendar error, go to:\nServices (+) in the left panel → Google Calendar API → Add", 8);
  } catch (e) {
    Logger.log("setupIncrementalSync error: " + e.message + "\n" + e.stack);
    sendTelegramError_("Setup incremental sync failed: " + e.message);
    throw e;
  }
}

function removeIncrementalSync() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "syncCalendarIncremental_") {
      ScriptApp.deleteTrigger(t);
    }
  });
  PropertiesService.getScriptProperties().deleteProperty("CALENDAR_SYNC_TOKEN");
  notify_("✅ Incremental sync stopped.");
}

/**
 * Installs just the midnight sweep trigger (if setupTriggers was already run).
 */
function setupMidnightSweep() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "runMidnightSweep_") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("runMidnightSweep_").timeBased().everyDays(1).atHour(0).create();
  notify_("✅ Midnight sweep enabled.");
}

/**
 * Runs at 00:00 daily — flips stale Upcoming rows to Not Paid without needing a calendar event.
 */
function runMidnightSweep_() {
  try {
    Logger.log("runMidnightSweep_ started");
    const ss = SpreadsheetApp.getActive();
    const sheet = getSheetOrThrow_(ss, APPOINTMENTS_SHEET);
    updateUpcomingToNotPaid_(sheet);
    sortAndHideAppointments_(sheet);
    Logger.log("runMidnightSweep_ complete");
  } catch (e) {
    Logger.log("runMidnightSweep_ error: " + e.message + "\n" + e.stack);
    sendTelegramError_("Midnight sweep failed: " + e.message);
  }
}

/**
 * Installs an installable onOpen trigger so the sheet syncs automatically when opened.
 */
function setupOnOpenSync() {
  try {
    Logger.log("setupOnOpenSync: installing onOpen trigger");
    ScriptApp.getProjectTriggers().forEach(t => {
      if (t.getHandlerFunction() === "onOpenSync_") ScriptApp.deleteTrigger(t);
    });
    ScriptApp.newTrigger("onOpenSync_").forSpreadsheet(SpreadsheetApp.getActive()).onOpen().create();
    Logger.log("setupOnOpenSync complete");
    notify_("✅ Sync on sheet open enabled.");
  } catch (e) {
    Logger.log("setupOnOpenSync error: " + e.message + "\n" + e.stack);
    sendTelegramError_("Setup onOpen sync failed: " + e.message);
    throw e;
  }
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
  lines.push(triggers.includes("runMidnightSweep_") ? "✅ Midnight sweep active" : "⚠️ Midnight sweep missing — run setupMidnightSweep()");
  lines.push(triggers.includes("sendDailyNotification") ? "✅ Daily notification trigger active" : "⚠️ Daily notification missing — run setupNotificationTrigger()");
  lines.push(triggers.includes("onOpenSync_") ? "✅ onOpen sync active" : "ℹ️ onOpen sync not set (optional — run 📲 Setup Sync on Sheet Open)");

  const syncToken = props.getProperty("CALENDAR_SYNC_TOKEN");
  lines.push(syncToken ? "✅ Calendar sync token present" : "ℹ️ No sync token yet (runs full sync on first trigger fire)");

  Logger.log("Setup validation:\n\n" + lines.join("\n"));
  notify_("Validation complete — see Logs (View → Logs) for full results.", 5);
}

/***************
 * MAINTENANCE
 ***************/

/**
 * One-time migration to populate the Cached Name column (L=12).
 * Run once from the Apps Script editor dropdown.
 */
function runMigration() {
  try {
    Logger.log("runMigration started");
    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName(APPOINTMENTS_SHEET);
    const lastRow = sheet.getLastRow();

    if (lastRow < DATA_ROW) { notify_("No appointment data to migrate."); return; }
    Logger.log(`Migrating ${lastRow - DATA_ROW + 1} rows`);

    // Ensure col L header exists (CachedName, col 12)
    if (!sheet.getRange(HEADER_ROW, 12).getValue()) {
      sheet.getRange(HEADER_ROW, 12).setValue("Cached Name");
    }

    // Read 13 cols from col B (B-N)
    const data = sheet.getRange(DATA_ROW, 2, lastRow - 2, 13).getValues();
    const cachedNames = [];
    const updatedFormulas = [];

    for (let i = 0; i < data.length; i++) {
      const row = i + DATA_ROW;
      const currentName = String(data[i][2] || "");   // index 2 = col D = Name
      const existingCache = String(data[i][10] || ""); // index 10 = col L = CachedName

      cachedNames.push([existingCache || currentName]);
      // M = ClientID, L = CachedName fallback
      updatedFormulas.push([`=XLOOKUP(M${row}; Clients!Q:Q; Clients!B:B; L${row})`]);
    }

    sheet.getRange(DATA_ROW, 12, cachedNames.length, 1).setValues(cachedNames);   // col L = CachedName
    sheet.getRange(DATA_ROW, 4, updatedFormulas.length, 1).setFormulas(updatedFormulas); // col D = Name formula

    Logger.log(`runMigration complete: ${cachedNames.length} rows updated`);
    notify_(`✅ Migration complete!\n${cachedNames.length} rows updated.\n\nYou can now hide columns L, M, N (right-click → Hide column).`, 8);
  } catch (e) {
    Logger.log("runMigration error: " + e.message + "\n" + e.stack);
    sendTelegramError_("Migration failed: " + e.message);
    throw e;
  }
}

/**
 * Removes duplicate rows safely.
 * Safe: only deletes rows where Payment is empty AND status is Not Paid/Upcoming
 * AND another row exists for the same client at the same date+time.
 */
function cleanupDuplicates() {
  try {
    Logger.log("cleanupDuplicates started");
    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName(APPOINTMENTS_SHEET);
    const lastRow = sheet.getLastRow();
    if (lastRow < DATA_ROW) { notify_("No data found."); return; }
    Logger.log(`Scanning ${lastRow - DATA_ROW + 1} rows for duplicates`);

    // Read 13 cols from col B (B-N)
    const data = sheet.getRange(DATA_ROW, 2, lastRow - 2, 13).getValues();

    const groups = new Map();
    for (let i = 0; i < data.length; i++) {
      const row     = data[i];
      const dateStr = row[0] ? new Date(row[0]).toDateString() : ""; // index 0 = col B = Date
      const timeStr = row[1] ? String(row[1]) : "";                  // index 1 = col C = Time
      const name    = String(row[2] || row[10] || "").trim().toLowerCase(); // col D=Name, col L=CachedName
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
      Logger.log("cleanupDuplicates: no duplicates found");
      notify_("✅ No duplicates found — sheet looks clean!");
      return;
    }

    Logger.log(`cleanupDuplicates: deleting ${toDelete.length} rows`);
    toDelete.sort((a, b) => b - a);
    for (const rowIdx of toDelete) {
      sheet.deleteRow(rowIdx);
    }

    Logger.log(`cleanupDuplicates complete: removed ${toDelete.length} rows`);
    notify_(`✅ Removed ${toDelete.length} duplicate rows.\n\nNow run: ✂️ Barber Tools → Setup Incremental Sync`);
  } catch (e) {
    Logger.log("cleanupDuplicates error: " + e.message + "\n" + e.stack);
    sendTelegramError_("Cleanup duplicates failed: " + e.message);
    throw e;
  }
}
