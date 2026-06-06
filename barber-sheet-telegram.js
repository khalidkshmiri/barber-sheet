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
  notify_("✅ Daily notification set for 9 PM (21:00) every evening.");
}

function removeNotificationTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "sendDailyNotification") ScriptApp.deleteTrigger(t);
  });
}

function sendDailyNotification() {
  sendDailyImageNotification_();
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
  if (lastRow < DATA_ROW) return;

  const data      = apptSheet.getRange(DATA_ROW, 2, lastRow - 2, 13).getValues();
  const clientMap = loadClientNotificationMap_(clientsSheet);
  const newAppts  = [];

  for (const row of data) {
    const eventId = String(row[12] || "");
    if (!newEventIds || !newEventIds.has(eventId)) continue;

    const name    = row[2] || row[10];
    const dateVal = row[0];
    const timeVal = row[1];
    const price   = row[3];
    const payment = row[4];
    const notes   = row[8];
    if (!name || !dateVal) continue;

    const dateStr  = Utilities.formatDate(new Date(dateVal), tz, "EEEE d MMMM");
    const timeStr  = timeVal ? Utilities.formatDate(new Date(timeVal), tz, "HH:mm") : "—";
    const priceStr = payment === "Subscription" ? "Sub" : "€" + price;

    const { badge, doNotCut } = getReliabilityInfo_(name, clientMap);
    const clientInfo          = clientMap.get(nameCase_(name)) || {};
    const vipBadge            = clientInfo.vip            ? " ⭐" : "";
    const consecutivePaid     = clientInfo.consecutivePaid || 0;
    const loyaltyLabel        = consecutivePaid >= 5                 ? " ✅ ELIGIBLE FOR FREE" : "";
    const lastAppt            = getClientLastAppointment_(name, apptSheet);
    const nameDisplay         = doNotCut ? "<u>" + name + "</u>" : name;
    const needsUpfrontPayment = doNotCut || badge.includes("⚠️"); // ⚠️ or ⛔

    newAppts.push({ name: nameDisplay, dateStr, timeStr, priceStr, badge, vipBadge, loyaltyLabel, lastAppt, notes, needsUpfrontPayment });
  }

  if (newAppts.length === 0) return;

  const plural = newAppts.length > 1 ? "s" : "";
  let lines = ["✂️ <b>New Appointment" + plural + " Added</b>"];

  for (const a of newAppts) {
    lines.push("");
    lines.push("<b>" + a.dateStr + " at " + a.timeStr + "</b>");
    lines.push("👤 " + a.name + a.vipBadge + " " + a.badge + a.loyaltyLabel);
    if (a.needsUpfrontPayment) lines.push("💳 Request payment upfront");
    lines.push("💶 " + a.priceStr);
    if (a.lastAppt) {
      const notesPart = a.lastAppt.notes ? " — \"" + a.lastAppt.notes + "\"" : "";
      lines.push("📋 Last: " + a.lastAppt.label + notesPart);
    } else {
      lines.push("📋 First visit");
    }
    if (a.notes) lines.push("📝 \"" + a.notes + "\"");
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
  if (lastRow < DATA_ROW) return { appts: [], clientMap: new Map(), tomorrow };

  const data = apptSheet.getRange(DATA_ROW, 2, lastRow - 2, 13).getValues();
  const clientMap = loadClientNotificationMap_(clientsSheet);

  const appts = [];
  for (const row of data) {
    const dateVal = row[0];
    if (!dateVal) continue;
    const d = new Date(dateVal);
    if (d < tomorrow || d > tomorrowEnd) continue;
    const status = row[5];
    if (status === "Cancelled" || status === "No Show") continue;

    const name = row[2] || row[10];
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
  if (lastRow < DATA_ROW) return [];

  const data = apptSheet.getRange(DATA_ROW, 2, lastRow - 2, 13).getValues();
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
      name: row[2] || row[10],
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
  if (lastRow < DATA_ROW) return [];

  const data = apptSheet.getRange(DATA_ROW, 2, lastRow - 2, 13).getValues();
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

    const name = row[2] || row[10];

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
  const data = apptSheet.getRange(DATA_ROW, 2, apptSheet.getLastRow() - 2, 13).getValues();
  const clientAppts = [];

  for (const row of data) {
    const dateVal = row[0];
    if (!dateVal) continue;
    const d = new Date(dateVal);
    if (d >= today) continue; // Only past appointments
    const name = row[2] || row[10];
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
  if (last.payment === "Free") {
    label = "Free";
  } else if (last.status === "Paid") {
    label = `Paid (€${last.price})`;
  } else if (last.status.startsWith("Free")) {
    label = last.status; // "Free - Family" etc
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
  const isNewClient = !clientMap.has(nameCase_(name)) || info.totalVisits === 0;

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
  } else if (isNewClient) {
    badge = "🆕 New client";
    badgeColor = "#0c5460";
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
 * TELEGRAM notification with all sections (kept for manual debugging)
 */
function sendTelegramNotification_() {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty("TELEGRAM_BOT_TOKEN");
  const chatId = props.getProperty("TELEGRAM_CHAT_ID");

  if (!token || !chatId) {
    Logger.log("⚠️ Telegram not configured. Go to: Extensions → Apps Script → Project Settings → Script Properties");
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
  const dncRecommendations = getDncRecommendations_(clientMap);

  // Only send if there's something to report
  if (appts.length === 0 && todayUnpaid.length === 0 && unreliable.length === 0 && dncRecommendations.length === 0) return;

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

      const clientInfo      = clientMap.get(nameCase_(appt.name)) || {};
      const vipBadge        = clientInfo.vip            ? " ⭐" : "";
      const consecutivePaid = clientInfo.consecutivePaid || 0;
      const loyaltyLabel = consecutivePaid >= 5 ? " ✅ ELIGIBLE FOR FREE" : "";

      msg += `<b>${timeStr} — ${nameDisplay}${vipBadge}</b>\n`;
      msg += `${priceStr} · ${badge}${loyaltyLabel}\n`;

      // Payment upfront reminder for unreliable clients
      if (badge.includes("⚠️") || badge.includes("⛔")) {
        msg += `💳 Request payment upfront\n`;
      }

      // Last appointment info
      if (appt.lastAppt) {
        msg += `📋 Last: ${appt.lastAppt.label}`;
        if (appt.lastAppt.notes) {
          msg += ` — "${appt.lastAppt.notes}"`;
        }
        msg += "\n";
      } else {
        msg += `🗂️ First visit\n`;
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
      msg += `  💳 Request payment upfront\n`;
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
      msg += `  💳 Request payment upfront\n`;
    }
  }

  // ── SECTION 4: DNC recommendations ──
  if (dncRecommendations.length > 0) {
    msg += `\n🚫 <b>Overweeg Do Not Cut (${dncRecommendations.length})</b>\n`;
    for (const r of dncRecommendations) {
      msg += `• ${r.name} — ${r.reason}\n`;
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

/**
 * Returns clients who have ≥2 no-shows OR ≥2 lates in the last 12 months but no DNC flag set.
 * These are suggestion-only — DNC is never auto-set.
 */
function getDncRecommendations_(clientMap) {
  const recs = [];
  clientMap.forEach((info, name) => {
    if (info.doNotCut) return; // already flagged
    const noShows = info.noShow || 0;
    const lates = info.late || 0;
    if (noShows < 2 && lates < 2) return;
    const parts = [];
    if (noShows >= 2) parts.push(`${noShows} no-shows`);
    if (lates >= 2) parts.push(`${lates} late`);
    recs.push({ name, reason: parts.join(', ') });
  });
  return recs;
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
      doNotCut: data[i][14] === true,
      vip: data[i][13] === true,
      consecutivePaid: Number(data[i][12]) || 0,
      totalVisits: Number(data[i][8]) || 0
    });
  }
  return map;
}

/**
 * IMAGE NOTIFICATION — builds card via Screenshotone and sends as photo to Telegram
 */
function sendDailyImageNotification_() {
  const cal = getCalendarOrThrow_();
  const tz = cal.getTimeZone();
  const { appts, clientMap, tomorrow } = getTomorrowAppointments_();
  const unpaid = getTodayUnpaid_();
  const unreliable = getUnreliableAppointments_(appts.map(a => nameCase_(a.name)));
  const dncRecs = getDncRecommendations_(clientMap);
  if (!appts.length && !unpaid.length && !unreliable.length && !dncRecs.length) return;
  const html = buildCardHtml_(appts, clientMap, unpaid, unreliable, dncRecs, tomorrow, tz);
  const imageBlob = generateCardImage_(html);
  if (imageBlob) sendTelegramPhoto_(imageBlob);
}

function buildCardHtml_(appts, clientMap, unpaid, unreliable, dncRecs, tomorrow, tz) {
  function esc(s) {
    return String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }
  function fmt(t) {
    if (!t || typeof t.getTime !== 'function') return '–';
    return Utilities.formatDate(new Date(t), tz, 'HH:mm');
  }
  function priceStr(appt) {
    return appt.payment === 'Subscription' ? 'Sub' : '€' + (appt.price || 0);
  }

  const dateRaw = Utilities.formatDate(tomorrow, tz, 'EEEE d MMMM');
  const dateStr = dateRaw.charAt(0).toUpperCase() + dateRaw.slice(1);

  let apptRows = '';
  for (const appt of appts) {
    const { badge, doNotCut } = getReliabilityInfo_(appt.name, clientMap);
    const info = clientMap.get(nameCase_(appt.name)) || {};
    const vip = info.vip ? ' <span class="badge">⭐</span>' : '';
    const free = (info.consecutivePaid || 0) >= 5 ? '<span class="badge gold">✅ FREE</span>' : '';
    const dnc = doNotCut ? ' do-not-cut' : '';
    apptRows += `<div class="row">
      <div class="pill">${esc(fmt(appt.time))}</div>
      <div class="info">
        <div class="name${dnc}">${esc(appt.name)}${vip}</div>
        <div class="sub">${esc(appt.service || '–')} · ${esc(priceStr(appt))}</div>
        <div class="badges"><span class="badge">${esc(badge)}</span>${free}</div>
      </div>
    </div>`;
  }

  let unpaidHtml = '';
  if (unpaid.length) {
    const rows = unpaid.map(u => `<div class="urow">
      <span class="usub">${esc(fmt(u.time))}</span>
      <span class="uname">${esc(u.name)}</span>
      <span class="uprice">€${u.price || 0}</span>
    </div>`).join('');
    unpaidHtml = `<div class="section">VANDAAG ONBETAALD</div>${rows}`;
  }

  let alertHtml = '';
  if (unreliable.length) {
    const rows = unreliable.map(u => {
      const dayLabel = u.daysAhead === 2 ? 'Overmorgen' : 'Morgen';
      const { badge: b2 } = getReliabilityInfo_(u.name, clientMap);
      return `<div class="urow">
        <span class="usub">${esc(dayLabel)}</span>
        <span class="uname">${esc(u.name)}</span>
        <span class="badge">${esc(b2)}</span>
      </div>`;
    }).join('');
    alertHtml = `<div class="section">LET OP</div>${rows}`;
  }

  let dncHtml = '';
  if (dncRecs && dncRecs.length) {
    const rows = dncRecs.map(r => `<div class="urow">
      <span class="usub">DNC?</span>
      <span class="uname">${esc(r.name)}</span>
      <span class="badge">${esc(r.reason)}</span>
    </div>`).join('');
    dncHtml = `<div class="section">OVERWEEG DO NOT CUT</div>${rows}`;
  }

  return `<!DOCTYPE html><html><head><meta charset="UTF-8">
<style>
*{margin:0;padding:0;box-sizing:border-box}
body{background:#15110C;font-family:'Inter','Helvetica Neue',Arial,sans-serif;width:390px}
.card{padding:24px 20px 28px}
.lbl{font-size:10px;font-weight:700;letter-spacing:2.5px;color:#E2B262;text-transform:uppercase;margin-bottom:4px}
.date{font-size:22px;font-weight:600;color:#F0E9DA;margin-bottom:18px}
.hr{height:1px;background:#2E2619;margin-bottom:16px}
.row{display:flex;align-items:flex-start;gap:12px;margin-bottom:14px}
.pill{background:#E2B262;color:#15110C;font-size:11px;font-weight:700;padding:3px 8px;border-radius:4px;white-space:nowrap;margin-top:2px;min-width:46px;text-align:center}
.info{flex:1;min-width:0}
.name{font-size:15px;font-weight:600;color:#F0E9DA;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.name.do-not-cut{text-decoration:line-through;color:#8A8070}
.sub{font-size:12px;color:#8A8070;margin-top:2px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.badges{display:flex;flex-wrap:wrap;gap:6px;margin-top:5px}
.badge{font-size:11px;color:#8A8070}
.badge.gold{color:#E2B262;font-weight:700}
.section{font-size:9px;font-weight:700;letter-spacing:2px;color:#8A8070;text-transform:uppercase;margin:16px 0 10px;border-top:1px solid #2E2619;padding-top:14px}
.urow{display:flex;align-items:center;gap:10px;margin-bottom:8px}
.usub{font-size:11px;color:#8A8070;min-width:56px}
.uname{font-size:13px;color:#F0E9DA;flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.uprice{font-size:13px;color:#E2B262}
.footer{margin-top:22px;font-size:9px;color:#2E2619;text-align:center;letter-spacing:1.5px;text-transform:uppercase}
</style></head><body>
<div id="card" class="card">
<div class="lbl">MORGEN</div>
<div class="date">${esc(dateStr)}</div>
<div class="hr"></div>
${apptRows}${unpaidHtml}${alertHtml}${dncHtml}
<div class="footer">Kashmir Barber · Rotterdam</div>
</div></body></html>`;
}

function generateCardImage_(html) {
  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty('SCREENSHOTONE_KEY');
  if (!apiKey) {
    Logger.log('SCREENSHOTONE_KEY missing in Script Properties');
    return null;
  }
  try {
    const resp = UrlFetchApp.fetch('https://api.screenshotone.com/take', {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        access_key: apiKey,
        html,
        format: 'jpg',
        image_quality: 90,
        viewport_width: 390,
        selector: '#card',
        device_scale_factor: 2
      }),
      muteHttpExceptions: true
    });
    const code = resp.getResponseCode();
    if (code !== 200) {
      Logger.log('Screenshotone error ' + code + ': ' + resp.getContentText());
      return null;
    }
    return resp.getBlob().setName('card.jpg');
  } catch (e) {
    Logger.log('Screenshotone error: ' + e.message);
    return null;
  }
}

function sendTelegramPhoto_(imageBlob) {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('TELEGRAM_BOT_TOKEN');
  const chatId = props.getProperty('TELEGRAM_CHAT_ID');
  if (!token || !chatId) return;
  try {
    UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/sendPhoto', {
      method: 'post',
      payload: { chat_id: chatId, photo: imageBlob }
    });
  } catch (e) {
    Logger.log('Telegram photo error: ' + e.message);
  }
}

function sendTelegramSimple_(text) {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty("TELEGRAM_BOT_TOKEN");
  const chatId = props.getProperty("TELEGRAM_CHAT_ID");
  if (!token || !chatId) return;
  try {
    UrlFetchApp.fetch("https://api.telegram.org/bot" + token + "/sendMessage", {
      method: "post", contentType: "application/json",
      payload: JSON.stringify({ chat_id: chatId, text })
    });
  } catch (e) { Logger.log("Telegram error: " + e.message); }
}

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
