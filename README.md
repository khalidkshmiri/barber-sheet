# barber-sheet

A Google Apps Script for managing a barber business — syncs appointments from Google Calendar (via iCloud/Apple Calendar), tracks clients, and sends Telegram notifications.

Built by Khalid (Rotterdam). Designed to work for anyone with a similar setup.

---

## What it does

- **Calendar sync** — pulls appointments from a Google Calendar (which can sync from Apple Calendar via iCloud)
- **Client records** — auto-creates client profiles, tracks visit history, reliability, VIP status, and consecutive paid streaks
- **Telegram notifications** — daily image card (9 PM), instant text alert when new appointments are synced
- **Subscription support** — tracks monthly subscription credits (up to 4 haircuts/month)
- **Smart triggers** — 5-minute incremental sync, midnight sweep, optional sync on sheet open, mobile-friendly Dashboard button
- **Dashboard period selector** — pick Today / This Week / This Month / Last Month / Last 3 Months / YTD / All Time to set a date range in C5–C6

---

## Prerequisites

- A Google account with Google Sheets and Google Calendar
- A Telegram bot (for notifications) — create one via [@BotFather](https://t.me/BotFather)
- A [Screenshotone](https://screenshotone.com) account (free, 100 images/month) — for the evening image card
- Optional: Apple Calendar synced to Google Calendar via iCloud

---

## Setup

### 1. Create the Google Sheet

Create a new Google Spreadsheet with **five sheets** (tabs) named exactly:

| Tab name | Purpose |
|----------|---------|
| `Appointments` | One row per appointment |
| `Clients` | One row per client |
| `Services` | Service name + price list |
| `Subscriptions` | Monthly subscription records |
| `Dashboard` | Sync button (C3 checkbox) and period selector (C4) |

> Row 1 and column A are spacers. Tables start at B2. Column order is hardcoded — changes require updating both the spreadsheet and the constants in the script.

**Appointments columns (B–N):**
`Date · Time · Name (formula) · Price · Payment · Status · Tips · Late · Notes · Service · CachedName (hidden) · ClientID (hidden) · EventID (hidden)`

**Clients columns (B–Q):**
`Name · FavService · LastVisit · SocialMedia · Notes · NoShow(12m) · Late(12m) · Referral (hidden) · TotalVisits · TotalSpent · TotalTips · FirstVisit · ConsecutivePaid · VIP · DoNotCut · ClientID (hidden)`

**Services columns (B–C):**
`ServiceName · Price`  (e.g. `Haircut · 15`)

**Subscriptions columns (B–J):**
`Name (formula) · Price · Type · Expiry · Credits (formula) · Status · Notes · StartDate · ClientID`

> **Locale note:** If your Google Sheets uses Dutch/European locale, formulas use semicolons (`;`) as separators instead of commas. The XLOOKUP formulas in the script already use semicolons. If you're on English locale, change them to commas in the script (lines with `=XLOOKUP`).

### 2. Name your Google Calendar

Create a Google Calendar named exactly `Barber Appointments` (or change `CALENDAR_NAME` at the top of the script).

Calendar event format: `ClientName - Service`  
Examples: `Ibrahim - Haircut`, `Youssef - Beard Trim`

### 3. Add the scripts

The project uses six script files that share the same global scope:

1. Open your Google Sheet → **Extensions → Apps Script**
2. Delete any existing code in the default file, paste the contents of `barber-sheet-script.js`, and rename it `barber-sheet-script`
3. Click **+** next to **Files** and add each of the remaining files the same way:

| File | Contents |
|------|----------|
| `barber-sheet-sync` | Calendar sync engines |
| `barber-sheet-clients` | Client stat recalculation |
| `barber-sheet-helpers` | Utility functions and index loaders |
| `barber-sheet-telegram` | Telegram notification logic |
| `barber-sheet-format` | Visual theme for all sheets |

### 4. Enable the Calendar API

In the Apps Script editor:
1. Click **Services** (the `+` icon in the left panel)
2. Find **Google Calendar API** → click **Add**

This is required for the incremental sync (5-minute trigger).

### 5. Set Script Properties

In the Apps Script editor:
1. Go to **Project Settings** (gear icon) → **Script Properties**
2. Add these properties:

| Property | Value |
|----------|-------|
| `TELEGRAM_BOT_TOKEN` | Your bot token from @BotFather (e.g. `123456:ABC-DEF...`) |
| `TELEGRAM_CHAT_ID` | Your Telegram user/chat ID (find it via @userinfobot) |
| `SCREENSHOTONE_KEY` | Your access key from [screenshotone.com](https://screenshotone.com) dashboard |

If you don't want Telegram notifications, leave the Telegram properties empty — the script will skip notifications. `SCREENSHOTONE_KEY` is only needed for the evening image card.

### 6. Run the setup

In the Apps Script editor, select `setupTriggers` from the function dropdown and click **Run**. Authorise the script when prompted.

This installs:
- An **onEdit trigger** (`processSheetChanges`) — handles the Dashboard sync button and sheet logic
- A **5-minute trigger** (`syncCalendarIncremental_`) — automatic calendar sync
- A **daily midnight trigger** (`runMidnightSweep_`) — flips stale Upcoming rows to Not Paid

### 7. Optional: sync when the sheet opens

Select `setupOnOpenSync` from the function dropdown and run it. This runs an incremental sync every time you open the spreadsheet.

### 8. Optional: daily image card at 9 PM

Select `setupNotificationTrigger` from the function dropdown and run it. This schedules a daily 9 PM trigger that generates an image card of tomorrow's appointments (via Screenshotone) and sends it to Telegram as a photo.

Requires `SCREENSHOTONE_KEY` to be set in Script Properties. To test it manually, select `sendDailyImageNotification_` from the dropdown and run it — note it only sends if there are appointments tomorrow, unpaid from today, unreliable upcoming clients, or DNC recommendations.

### 9. Validate your setup

Select `validateSetup` from the function dropdown and run it. Check **View → Logs** for the full results — it confirms the calendar, sheets, credentials, and triggers are all working.

---

## Optional: iOS Shortcut (instant sync)

Deploy the script as a web app:
1. Apps Script editor → **Deploy → Manage deployments → New deployment**
2. Type: **Web app**, execute as: **Me**, access: **Anyone**
3. Copy the web app URL

Create an iOS Shortcut that calls this URL when the Calendar app closes — this triggers an immediate incremental sync.

---

## Configuration

All settings are at the top of `barber-sheet-script.js`.

```js
const CALENDAR_NAME = "Barber Appointments";  // must match your Google Calendar name
const APPOINTMENTS_SHEET = "Appointments";    // sheet tab names
const CLIENTS_SHEET = "Clients";
const SERVICES_SHEET = "Services";
const SUBSCRIPTIONS_SHEET = "Subscriptions";

const DAYS_BACK = 14;            // how far back to sync
const DAYS_FORWARD = 14;         // how far forward to sync
const HIDE_OLDER_THAN_DAYS = 30; // hide (not delete) rows older than this

const VIP_MIN_VISITS = 15;  // total paid visits to auto-earn VIP
const VIP_MIN_SPENT  = 400; // total € spent to auto-earn VIP
```

VIP is auto-promoted when either threshold is met and is never auto-demoted. You can also set it manually via the checkbox in the Clients sheet.

---

## Payment types (col F)

| Value | Effect |
|-------|--------|
| `Cash` | Sets status → Paid |
| `Tikkie` | Sets status → Paid |
| `Subscription` | Sets status → Paid, price → €0 |
| `Free` | Sets status → Paid, price → €0, counts toward loyalty streak |

## Status values (col G)

| Value | Notes |
|-------|-------|
| `Upcoming` | Future appointment, not yet paid |
| `Not Paid` | Past appointment, no payment recorded |
| `Paid` | Payment received |
| `No Show` | Client didn't show up; clears Late checkbox |
| `Cancelled` | Appointment cancelled; clears Late checkbox |

---

## Reliability badges

Shown in Telegram notifications based on combined No Show + Late count (last 12 months):

| Badge | Condition |
|-------|-----------|
| ⛔ DO NOT CUT | DoNotCut checkbox is set |
| ⚠️ Unreliable | 3+ incidents |
| 🟡 Watch | 1–2 incidents |
| 🆕 New client | No prior visits on record |
| ✅ Reliable | 0 incidents |

Clients with 5+ consecutive paid appointments without a no-show or late are flagged **✅ ELIGIBLE FOR FREE** in notifications.

The daily notification also includes a **DNC recommendations** section for clients with ≥2 no-shows or ≥2 lates who don't have the DoNotCut flag set yet. These are suggestions only — the flag is never auto-set.

---

## Useful manual functions

Run these from the **function dropdown** in the Apps Script editor:

| Function | What it does |
|----------|-------------|
| `syncThisYear` | Full sync for the entire current year |
| `setupTriggers` | Install onEdit + 5-min sync + midnight sweep triggers |
| `setupIncrementalSync` | Install just the 5-min incremental sync trigger |
| `removeIncrementalSync` | Remove the 5-min sync trigger |
| `setupMidnightSweep` | Install just the daily midnight sweep trigger |
| `setupNotificationTrigger` | Schedule daily 9 PM Telegram notification |
| `removeNotificationTrigger` | Remove the daily notification |
| `setupOnOpenSync` | Sync on every sheet open |
| `validateSetup` | Check everything is configured correctly |
| `formatSpreadsheet` | Apply the visual theme to all sheets (safe to re-run) |
| `runMigration` | One-time: populate CachedName column (col L) |
| `cleanupDuplicates` | Remove duplicate appointment rows |
| `testTelegramAPI` | Test the Telegram bot connection |
| `debugTomorrow` | Log tomorrow's appointments to the console |
