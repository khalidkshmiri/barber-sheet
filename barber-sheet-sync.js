/***************
 * SYNC ENGINES
 * Full sync, year sync, incremental sync, and the upsert/sort/hide logic.
 ***************/

function syncCalendarToSheets(showNotification = true) {
  try {
    Logger.log("syncCalendarToSheets started");
    const ctx = prepareContext_();
    const now = new Date();
    const startDate = startOfDay_(addDays_(now, -DAYS_BACK));
    const endDate = endOfDay_(addDays_(now, DAYS_FORWARD));
    Logger.log(`Fetching events from ${startDate.toDateString()} to ${endDate.toDateString()}`);
    const events = ctx.calendar.getEvents(startDate, endDate);
    Logger.log(`Found ${events.length} calendar events`);

    const counts = upsertEvents_(events, { ...ctx, startDate, endDate });
    Logger.log(`Upsert done: +${counts.newCount} new, ~${counts.updatedCount} updated, -${counts.cancelledCount} cancelled`);

    updateUpcomingToNotPaid_(ctx.appointmentsSheet);
    const apptLastRow = ctx.appointmentsSheet.getLastRow();
    const apptData = apptLastRow >= DATA_ROW ? ctx.appointmentsSheet.getRange(DATA_ROW, 2, apptLastRow - 2, 13).getValues() : [];
    updateConsecutivePaidCounts_(ctx, apptData);
    updateNoShowLateCounts_(ctx, apptData);
    updateClientStats_(ctx, apptData);
    sortAndHideAppointments_(ctx.appointmentsSheet);
    Logger.log("syncCalendarToSheets complete");

    if (showNotification) {
      notify_(`Sync Complete!\n\n+ ${counts.newCount} New\n~ ${counts.updatedCount} Updated\n- ${counts.cancelledCount} Cancelled`);
      if (counts.newCount > 0 || counts.cancelledCount > 0) {
        sendSyncNotification_(ctx, counts.newEventIds);
      }
    }
  } catch (e) {
    Logger.log("syncCalendarToSheets error: " + e.message + "\n" + e.stack);
    sendTelegramError_("Full sync failed: " + e.message);
    throw e;
  }
}

function syncThisYear(showNotification = true) {
  try {
    Logger.log("syncThisYear started");
    const ctx = prepareContext_();
    const now = new Date();
    const year = now.getFullYear();
    const startDate = new Date(year, 0, 1);
    const endDate = new Date(year, 11, 31, 23, 59, 59, 999);
    Logger.log(`Fetching all events for ${year}`);

    const events = ctx.calendar.getEvents(startDate, endDate);
    Logger.log(`Found ${events.length} calendar events`);

    const counts = upsertEvents_(events, { ...ctx, startDate, endDate });
    Logger.log(`Upsert done: +${counts.newCount} new, ~${counts.updatedCount} updated, -${counts.cancelledCount} cancelled`);

    updateUpcomingToNotPaid_(ctx.appointmentsSheet);
    const apptLastRowY = ctx.appointmentsSheet.getLastRow();
    const apptDataY = apptLastRowY >= DATA_ROW ? ctx.appointmentsSheet.getRange(DATA_ROW, 2, apptLastRowY - 2, 13).getValues() : [];
    updateConsecutivePaidCounts_(ctx, apptDataY);
    updateNoShowLateCounts_(ctx, apptDataY);
    updateClientStats_(ctx, apptDataY);
    sortAndHideAppointments_(ctx.appointmentsSheet);
    Logger.log("syncThisYear complete");

    if (showNotification) {
      notify_(`Sync Complete!\n\n+ ${counts.newCount} New\n~ ${counts.updatedCount} Updated\n- ${counts.cancelledCount} Cancelled`);
      if (counts.newCount > 0 || counts.cancelledCount > 0) {
        sendSyncNotification_(ctx, counts.newEventIds);
      }
    }
  } catch (e) {
    Logger.log("syncThisYear error: " + e.message + "\n" + e.stack);
    sendTelegramError_("Year sync failed: " + e.message);
    throw e;
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
        event.getDescription() || "", serviceToWrite, parsed.clientName, clientId, eventId
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

        // col D=4 Name formula; M=ClientID, L=CachedName fallback
        const nameFormula = `=XLOOKUP(M${existingRow}; Clients!Q:Q; Clients!B:B; L${existingRow})`;
        ctx.appointmentsSheet.getRange(existingRow, 2, 1, 3).setValues([[dateCell, timeCell, nameFormula]]); // cols B-D
        ctx.appointmentsSheet.getRange(existingRow, 5).setValue(isSubscriptionSale ? 0 : rowPrice);          // col E = Price
        ctx.appointmentsSheet.getRange(existingRow, 10, 1, 5).setValues([[                                   // cols J-N
          event.getDescription() || "", serviceToWrite, parsed.clientName, clientId, eventId
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
      // M = ClientID, L = CachedName fallback
      newRowsABC[i][2] = `=XLOOKUP(M${targetRow}; Clients!Q:Q; Clients!B:B; L${targetRow})`;
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

function sortAndHideAppointments_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_ROW) return;

  // Sort data rows (starting at DATA_ROW, from col B). Column numbers are relative to the range.
  // Column 1 of range = col B = Date, Column 2 = col C = Time.
  const numRows = lastRow - DATA_ROW + 1;
  const numCols = Math.max(1, sheet.getLastColumn() - 1); // cols B onward
  sheet.getRange(DATA_ROW, 2, numRows, numCols).sort([
    { column: 2, ascending: false }, // col B = Date (absolute column number)
    { column: 3, ascending: false }  // col C = Time (absolute column number)
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

  if (ctx.subsIndex.byName.has(name) || (clientId && ctx.subsIndex.byId.has(String(clientId)))) {
    Logger.log(`createSubscriptionEntry_: skipping ${name} — already has active subscription`);
    return;
  }
  Logger.log(`createSubscriptionEntry_: creating subscription for ${name} (id=${clientId})`);

  const lastRow = subsSheet.getLastRow();
  if (lastRow >= DATA_ROW) {
    // Check last ~20 rows to avoid duplicate subscription entries
    const checkStart = Math.max(DATA_ROW, lastRow - 20);
    const checkCount = lastRow - checkStart + 1;
    const checkRange = subsSheet.getRange(checkStart, 2, checkCount, 8).getValues();
    const startYMD = ymd_(startDate instanceof Date ? startDate : new Date(startDate), ctx.calTz);
    for (const r of checkRange) {
      // index 7 = col I = StartDate
      if (nameCase_(r[0]) === name && r[7] && ymd_(new Date(r[7]), ctx.calTz) === startYMD) return;
    }
  }

  const r = lastRow + 1;
  // ClientID is in col J of Subscriptions (index 8 of the 9-col data, col B=Name through J=ClientID)
  // Clients!Q:Q = ClientID, Clients!B:B = Name
  const nameFormula = `=XLOOKUP(J${r}; Clients!Q:Q; Clients!B:B; "")`;
  // Subscriptions layout (starting col B): B=Name C=Price D=Type E=Expiry F=Credits G=Status H=Notes I=StartDate J=ClientID
  // I${r}=StartDate, E${r}=Expiry; Appointments cols: $M:$M=ClientID, $F:$F=Payment, $G:$G=Status, $B:$B=Date
  const creditsFormula = `=MAX(0; 4 - COUNTIFS(Appointments!$M:$M; J${r}; Appointments!$F:$F; "Subscription"; Appointments!$G:$G; "Paid"; Appointments!$B:$B; ">="&I${r}; Appointments!$B:$B; "<="&(E${r} + 21)))`;

  // Write to cols B-J: Name, Price, Type, Expiry, Credits, Status, Notes, StartDate, ClientID
  subsSheet.getRange(r, 2, 1, 9).setValues([[nameFormula, monthlyPrice, "", addDays_(startDate, 31), creditsFormula, "Active", "", startDate, clientId]]);

  const newLastRow = subsSheet.getLastRow();
  if (newLastRow >= DATA_ROW) {
    // Sort by StartDate (col I = absolute column 9)
    subsSheet.getRange(DATA_ROW, 2, newLastRow - 2, 9).sort({ column: 9, ascending: false });
  }

  const entry = { start: startDate };
  ctx.subsIndex.byName.set(name, entry);
  if (clientId) ctx.subsIndex.byId.set(String(clientId), entry);
}

/**
 * INCREMENTAL SYNC
 * SETUP: Apps Script editor → Services (+) → Google Calendar API → Add
 * Then run: setupTriggers() or setupIncrementalSync()
 */
function syncCalendarIncremental_() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(0)) {
    Logger.log("Sync already running, skipping.");
    return;
  }
  try {
    Logger.log("Incremental sync started");
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
    Logger.log("Items to process: " + items.length);
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
          item.description || "", serviceToWrite, parsed.clientName, clientId, eventId]);
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
          const formula = `=XLOOKUP(M${existingRow}; Clients!Q:Q; Clients!B:B; L${existingRow})`;
          apptSheet.getRange(existingRow, 2, 1, 3).setValues([[dateCell, timeCell, formula]]); // cols B-D
          apptSheet.getRange(existingRow, 10, 1, 5).setValues([[                               // cols J-N
            item.description || "", serviceToWrite, parsed.clientName, clientId, eventId
          ]]);
        }
      }
    }

    // Write new rows
    if (newRowsABC.length > 0) {
      const startRow = apptSheet.getLastRow() + 1;
      for (let i = 0; i < newRowsABC.length; i++) {
        const r = startRow + i;
        newRowsABC[i][2] = `=XLOOKUP(M${r}; Clients!Q:Q; Clients!B:B; L${r})`;
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
    Logger.log("Incremental sync complete");

  } catch (e) {
    Logger.log("Incremental sync error: " + e.message + "\n" + e.stack);
    sendTelegramError_("Sync failed: " + e.message);
  } finally {
    lock.releaseLock();
  }
}
