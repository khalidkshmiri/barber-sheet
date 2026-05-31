# barber-sheet

A Google Apps Script for managing a barber business — syncs appointments from Google Calendar (via iCloud/Apple Calendar), tracks clients, and sends Telegram notifications.

Built by Khalid (Rotterdam). Designed to work for anyone with a similar setup.

---

## What it does

- **Calendar sync** — pulls appointments from a Google Calendar (which can sync from Apple Calendar via iCloud)
- **Client records** — auto-creates client profiles, tracks visit history, reliability, VIP status, and consecutive paid streaks
- **Telegram notifications** — daily summary (9 PM), instant alert when new appointments are synced
- **Subscription support** — tracks monthly subscription credits (up to 4 haircuts/month)
- **Smart triggers** — 5-minute incremental sync, optional sync on sheet open, mobile-friendly Dashboard button

---

## Prerequisites

- A Google account with Google Sheets and Google Calendar
- A Telegram bot (for notifications) — create one via [@BotFather](https://t.me/BotFather)
- Optional: Apple Calendar synced to Google Calendar via iCloud

---

## Setup

### 1. Create the Google Sheet

Create a new Google Spreadsheet with **four sheets** (tabs) named exactly:

| Tab name | Purpose |
|----------|---------|
| `Appointments` | One row per appointment |
| `Clients` | One row per client |
| `Services` | Service name + price list |
| `Subscriptions` | Monthly subscription records |

**Appointments columns (A–M):**
`Date | Time | Name (formula) | Price | Payment | Status | Tips | Late | Notes | Service | ClientID | EventID | CachedName`

**Clients columns (A–P):**
`Name | FavService | LastVisit | SocialMedia | Notes | NoShow(12m) | Late(12m) | Referral | TotalVisits | TotalTips | TotalSpent | ClientID | FirstVisit | DoNotCut | ConsecutivePaid | VIP`

**Services columns:**
`ServiceName | Price`  (e.g. `Haircut | 15`)

**Subscriptions columns (A–I):**
`Name (formula) | Price | Notes | StartDate | Credits (formula) | Status | ExpiryDate | Notes | ClientID`

> **Locale note:** If your Google Sheets uses Dutch/European locale, formulas use semicolons (`;`) as separators instead of commas. The XLOOKUP formulas in the script already use semicolons. If you're on English locale, change them to commas in the script (lines with `=XLOOKUP`).

### 2. Name your Google Calendar

Create a Google Calendar named exactly `Barber Appointments` (or change `CALENDAR_NAME` at the top of the script).

Calendar event format: `ClientName - Service`  
Examples: `Ibrahim - Haircut`, `Youssef - Beard Trim`

### 3. Add the script

1. Open your Google Sheet
2. Go to **Extensions → Apps Script**
3. Delete any existing code and paste the full contents of `barber-sheet-script.js`
4. Click **Save**

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

If you don't want Telegram notifications, leave these empty — the script will just skip notifications.

### 6. Run the setup

1. Back in your Google Sheet, reload the page
2. Click **✂️ Barber Tools → 🛠️ Setup All Triggers**
3. Authorize the script when prompted (Google will show a permissions screen)

This installs:
- An **onEdit trigger** (`processSheetChanges`) for the Dashboard button and sheet logic
- A **5-minute trigger** (`syncCalendarIncremental_`) for automatic calendar sync

### 7. Optional: sync when the sheet opens

Click **✂️ Barber Tools → 📲 Setup Sync on Sheet Open** to run an incremental sync every time you open the spreadsheet on mobile.

### 8. Optional: daily notification at 9 PM

Run `setupNotificationTrigger()` from the Apps Script function dropdown to schedule the daily Telegram summary.

### 9. Validate your setup

Click **✂️ Barber Tools → ✅ Validate Setup** to confirm the calendar, sheets, credentials, and triggers are all working.

---

## Optional: iOS Shortcut (instant sync)

Deploy the script as a web app:
1. Apps Script editor → **Deploy → Manage deployments → New deployment**
2. Type: **Web app**, execute as: **Me**, access: **Anyone**
3. Copy the web app URL

Create an iOS Shortcut that calls this URL when the Calendar app closes — this triggers an immediate incremental sync.

---

## Configuration

All settings are at the top of `barber-sheet-script.js`:

```js
const CALENDAR_NAME = "Barber Appointments";  // must match your Google Calendar name
const APPOINTMENTS_SHEET = "Appointments";    // sheet tab names
const CLIENTS_SHEET = "Clients";
const SERVICES_SHEET = "Services";
const SUBSCRIPTIONS_SHEET = "Subscriptions";

const DAYS_BACK = 14;          // how far back to sync
const DAYS_FORWARD = 14;       // how far forward to sync
const HIDE_OLDER_THAN_DAYS = 30; // hide (not delete) rows older than this

const NOTIFICATION_MODE = "telegram"; // currently only "telegram" is supported
```

---

## Payment types (col E)

| Value | Effect |
|-------|--------|
| `Cash` | Sets status → Paid |
| `Tikkie` | Sets status → Paid |
| `Subscription` | Sets status → Paid, price → €0 |
| `Free` | Sets status → Paid, price → €0, counts toward loyalty streak |

## Status values (col F)

| Value | Notes |
|-------|-------|
| `Upcoming` | Future appointment, not yet paid |
| `Not Paid` | Past appointment, no payment recorded |
| `Paid` | Payment received |
| `No Show` | Client didn't show up; clears Late checkbox |
| `Cancelled` | Appointment cancelled; clears Late checkbox |
| `Free` | Legacy: also sets price → €0 (use Payment = Free instead) |

---

## Useful manual functions

Run these from the **function dropdown** in the Apps Script editor:

| Function | What it does |
|----------|-------------|
| `syncThisYear` | Full sync for the entire current year |
| `setupNotificationTrigger` | Schedule daily 9 PM Telegram notification |
| `removeNotificationTrigger` | Remove the daily notification |
| `setupTriggers` | Install onEdit + 5-min sync triggers |
| `setupOnOpenSync` | Sync on every sheet open |
| `validateSetup` | Check everything is configured correctly |
| `runMigration` | One-time: populate CachedName column (col M) |
| `cleanupDuplicates` | Remove duplicate appointment rows |
| `testTelegramAPI` | Test the Telegram bot connection |
| `debugTomorrow` | Log tomorrow's appointments to the console |
