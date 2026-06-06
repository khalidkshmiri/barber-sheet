/***************
 * CLIENT STATS
 * Recalculates derived columns in the Clients sheet from appointment data.
 ***************/

/**
 * Update Consecutive Paid counts in Clients sheet (col N=14)
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
    const clientId = clientData[i][15]; // index 15 = col Q = ClientID
    if (!clientId) { oVals[i] = [0]; continue; }

    const clientAppts = [];
    for (const row of apptData) {
      // index 11 = col M = ClientID in appointments
      if (String(row[11]) === String(clientId)) {
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

  clientsSheet.getRange(DATA_ROW, 14, n, 1).setValues(oVals); // col N = ConsecutivePaid
}

/**
 * Update TotalVisits (J=10), TotalSpent (K=11), TotalTips (L=12), LastVisit (D=4), FirstVisit (M=13),
 * and VIP (O=15) in Clients sheet.
 */
function updateClientStats_(ctx, apptData) {
  const clientsSheet = ctx.clientsSheet;
  const apptSheet = ctx.appointmentsSheet;
  const clientLastRow = clientsSheet.getLastRow();
  if (clientLastRow < DATA_ROW) return;

  // Read 16 cols starting at col B (cols B-Q) to reach ClientID at Q
  const clientData = clientsSheet.getRange(DATA_ROW, 2, clientLastRow - 2, 16).getValues();
  if (!apptData) {
    const apptLastRow = apptSheet.getLastRow();
    apptData = apptLastRow >= DATA_ROW ? apptSheet.getRange(DATA_ROW, 2, apptLastRow - 2, 12).getValues() : [];
  }

  const lastVisits = [];   // col D = LastVisit
  const ijkValues = [];    // cols J, K, L (TotalVisits, TotalSpent, TotalTips)
  const firstVisits = [];  // col M = FirstVisit
  const vipValues = [];    // col O = VIP (auto-promoted, never demoted)

  for (let i = 0; i < clientData.length; i++) {
    const clientId = clientData[i][15]; // index 15 = col Q = ClientID
    if (!clientId) {
      lastVisits.push([""]);
      ijkValues.push([0, 0, 0]);
      firstVisits.push([""]);
      vipValues.push([clientData[i][13] === true]); // preserve existing VIP flag
      continue;
    }

    let totalVisits = 0, totalTips = 0, totalSpent = 0;
    let firstVisit = null, lastVisit = null;

    for (const row of apptData) {
      if (String(row[11]) !== String(clientId)) continue; // index 11 = col M = ClientID

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
    ijkValues.push([totalVisits, totalSpent, totalTips]); // K=TotalSpent, L=TotalTips
    firstVisits.push([firstVisit || ""]);

    // Auto-promote to VIP if thresholds met; never demote existing true
    const currentVip = clientData[i][13]; // index 13 = col O = VIP (16 cols read from B)
    const earnedVip  = totalVisits >= VIP_MIN_VISITS || totalSpent >= VIP_MIN_SPENT;
    vipValues.push([currentVip === true || earnedVip]);
  }

  const n = clientData.length;
  clientsSheet.getRange(DATA_ROW, 4,  n, 1).setValues(lastVisits);  // col D = LastVisit
  clientsSheet.getRange(DATA_ROW, 10, n, 3).setValues(ijkValues);   // cols J, K, L
  clientsSheet.getRange(DATA_ROW, 13, n, 1).setValues(firstVisits); // col M = FirstVisit
  clientsSheet.getRange(DATA_ROW, 15, n, 1).setValues(vipValues);   // col O = VIP
}

/**
 * Update NoShow (col G=7) and Late (col H=8) counts in Clients sheet.
 * Only counts incidents within the last 12 months.
 */
function updateNoShowLateCounts_(ctx, apptData) {
  const clientsSheet = ctx.clientsSheet;
  const apptSheet = ctx.appointmentsSheet;
  const clientLastRow = clientsSheet.getLastRow();
  if (clientLastRow < DATA_ROW) return;

  if (!apptData) {
    const apptLastRow = apptSheet.getLastRow();
    if (apptLastRow < DATA_ROW) return;
    apptData = apptSheet.getRange(DATA_ROW, 2, apptLastRow - 2, 12).getValues();
  }

  const cutoff = new Date();
  cutoff.setFullYear(cutoff.getFullYear() - 1);

  // Read 16 cols from col B (B-Q) to reach ClientID at Q
  const clientData = clientsSheet.getRange(DATA_ROW, 2, clientLastRow - 2, 16).getValues();
  const n = clientData.length;
  const fgVals = new Array(n);

  for (let i = 0; i < n; i++) {
    const clientId = clientData[i][15]; // index 15 = col Q = ClientID
    if (!clientId) { fgVals[i] = [0, 0]; continue; }

    let noShows = 0, lates = 0;
    for (const row of apptData) {
      if (String(row[11]) !== String(clientId)) continue; // index 11 = col M = ClientID
      const dateVal = row[0];
      if (!dateVal || new Date(dateVal) < cutoff) continue;
      if (row[5] === "No Show") noShows++; // index 5 = col G = Status
      if (row[7] === true) lates++;        // index 7 = col I = Late
    }
    fgVals[i] = [noShows, lates];
  }

  clientsSheet.getRange(DATA_ROW, 7, n, 2).setValues(fgVals); // cols G-H = NoShow, Late
}
