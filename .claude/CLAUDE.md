# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What This Is

A Google Apps Script for a barber business management system. The script is split across 6 files that run inside a Google Sheets spreadsheet and sync appointments from Google Calendar (which syncs from Apple Calendar via iCloud), manage client records, and send Telegram notifications.

**Script files (each maps 1:1 to a file in the Apps Script editor, shown there with a `.gs` suffix — kept locally as `.js` and synced via `clasp`):**

- `barber-sheet-script.js` — main/core logic
- `barber-sheet-clients.js` — client management
- `barber-sheet-format.js` — formatting/theming
- `barber-sheet-helpers.js` — shared helper functions
- `barber-sheet-sync.js` — calendar sync logic
- `barber-sheet-telegram.js` — Telegram notifications

**Owner:** Khalid — barber based in Rotterdam, Netherlands

**Daily use:** The spreadsheet is used daily via the iOS Google Sheets app. Features, UI decisions, and interactions must account for a mobile-first experience.

**Mobile zoom level:** The sheet is used at **25% zoom (maximum zoom-out) on iPhone 16 Pro Max**. All sizing decisions must be based on this. At 25% zoom: visible sheet width ≈ 1720px, visible sheet height ≈ 3000px. Design target: ~10 data rows visible at once.

**Row heights:** Header row = 150px. Data rows = 260px (`setRowHeightsForced(..., 260)`). 26px, 44px, and 64px are all confirmed too small at 25% zoom.

**Column widths:** Visible columns in Appointments (A–J) should total ~1720px. Visible columns in Clients should also total ~1720px across the 9 non-hidden columns.

## Deployment

There are no build or test commands. The repo is linked to the Apps Script project via `clasp` (`.clasp.json` holds the script ID; `appsscript.json` is the manifest). Script files and `appsscript.json` live together in `scripts/` (set as `rootDir` in `.clasp.json`) — **never move `appsscript.json` out of `scripts/`**: clasp only finds and pushes the manifest when it sits inside `rootDir`, so moving it elsewhere silently breaks manifest syncing (timezone, OAuth scopes, web app config, etc. would drift from production without any error). Development workflow:

1. Edit the relevant `.js` file(s) locally (see file list above)
2. Run `clasp push` to sync local files to the Apps Script project — this uploads ALL tracked files in one shot (Apps Script doesn't support partial pushes; `clasp push` always overwrites the whole remote project with the local files, so there's no need to track "which file changed")
3. **Claude should run `clasp push` automatically after making script edits** — no need to ask Khalid to paste files manually anymore
4. To deploy as web app: **Deploy → Manage deployments → Edit → New version → Deploy** (still manual — `clasp push` updates script content but not a versioned web app deployment)
5. To run functions manually: select from the function dropdown in the editor and click Run, or `clasp run <functionName>` (requires the project to use OAuth scopes set up for running)

`clasp open-script` opens the project in the Apps Script editor in a browser. If `clasp push` fails with an auth error, run `clasp login` again.

Syntax checking: `node --check <file>.js`

After making any code change, run `clasp push` to sync it to the Apps Script project — Khalid no longer needs to paste files manually.

## Google Sheet Structure

Four sheets: **Appointments**, **Clients**, **Services**, **Subscriptions**

**Column layout:** Row 1 and column A are spacers. Last column in each tab is also a spacer. Tables start at B2. Column order is hardcoded — changes require updating both the spreadsheet and the constants in the script.

**Appointments columns (B=2 start):**
`Date` · `Time` · `Name` · `Price` · `Payment` · `Status` · `Tips` · `Late` · `Notes` · `Service` · `Cached Name` · `ClientID` · `EventID` · _(spacer)_

Hidden: `Cached Name` (L=12), `ClientID` (M=13), `EventID` (N=14)

**Clients columns (B=2 start):**
`Name` · `Favourite Service` · `Last Visit` · `Social Media` · `Notes` · `No Show` · `Late` · `Referral` · `Total Visits` · `Total Spent` · `Total Tips` · `First Visit` · `Consecutive Paid` · `VIP` · `Do Not Cut` · `ClientID` · _(spacer)_

Hidden: `Referral` (I=9), `ClientID` (Q=17)

**Services columns:** _(spacer)_ · `Service` · `Price` · _(spacer)_

**Subscriptions columns (B=2 start):**
`Name` · `Price` · `Type` · `Expiry` · `Credits` · `Status` · `Notes` · `Start Date` · `ClientID` · _(spacer)_

**Locale:** Dutch — Google Sheets formulas use **semicolons** as separators, not commas. Example: `=XLOOKUP(M3; Clients!Q:Q; Clients!B:B; L3)`

## Architecture

### Calendar Sync

Calendar event title format: `ClientName - Service` (e.g. `Ibrahim - Haircut`)

**iCalUID vs opaque ID (critical):** `CalendarApp.event.getId()` returns iCalUID (`abc123@google.com`). The Calendar REST API's `item.id` returns an opaque ID (`abc123`). These differ for the same event. The sheet always stores iCalUID in the `EventID` column (N=14). Incremental sync must use `item.iCalUID` (not `item.id`) — this was a bug that caused duplicate rows; it is fixed.

**Three sync paths:**

- `syncCalendarToSheets()` — manual full sync, ±14 days window, ~18 sec
- `syncCalendarIncremental_()` — 5-min trigger using Calendar API sync token; exits in ~0.4 sec if nothing changed; uses `LockService` to prevent concurrent runs
- `doGet(e)` — web app endpoint called by iOS Shortcut when Calendar app closes; calls `syncCalendarIncremental_()` and shares its sync token

After any sync: `updateUpcomingToNotPaid_()`, `updateConsecutivePaidCounts_()`, `updateNoShowLateCounts_()`, `sortAndHideAppointments_()` always run.

### Subscription Logic

When service is "Monthly Subscription": sets `serviceToWrite = "Haircut"`, price = 0, payment = "Subscription", and calls `createSubscriptionEntry_()`. A client with active subscription credits also gets price = 0, payment = "Subscription" automatically. Credits formula in Subscriptions sheet counts used Subscription-paid appointments since start date (max 4 per month).

### onEdit Trigger (`processSheetChanges`)

This is an **installable trigger** (not a simple `onEdit`). It handles:

- Dashboard C3 checkbox → triggers `syncCalendarIncremental_()` + sends a Telegram confirmation
- Appointments `Payment` change → auto-sets Status and Price. Valid payment values: Cash, Tikkie, Subscription, Free
- Appointments `Status` change → No Show/Cancelled clears Late checkbox; Free-\* sets price to €0
- Clients `Name` column → auto-formats name to Title Case

**Payment = "Free"** sets price → 0, status → "Paid", and counts toward the consecutive paid streak.

### Telegram Notifications

Script Properties store `TELEGRAM_BOT_TOKEN` and `TELEGRAM_CHAT_ID`. Two notification functions:

- `sendDailyNotification()` — called by 9 PM daily trigger; shows tomorrow's appointments, today's unpaid, unreliable client alerts
- `sendSyncNotification_(ctx, newEventIds)` — fires immediately when new appointments are added during sync

Reliability badges based on combined `No Show` + `Late` count from Clients sheet: ⛔ DO NOT CUT (`Do Not Cut` checkbox), ⚠️ Unreliable (3+), 🟡 Watch (1–2), ✅ Reliable (0).

Consecutive paid streak (`Consecutive Paid`) resets on No Show or Late checkbox; at 5+ shows "✅ ELIGIBLE FOR FREE" in notifications.

## Key Functions Reference

| Function                            | Purpose                                                              |
| ----------------------------------- | -------------------------------------------------------------------- |
| `upsertEvents_(events, ctx)`        | Core sync logic — processes calendar events into sheet rows          |
| `prepareContext_()`                 | Loads all sheet references and indexes into a context object         |
| `syncCalendarIncremental_()`        | Incremental sync using Calendar API sync token                       |
| `processSheetChanges(e)`            | Installable onEdit handler                                           |
| `sendTelegramNotification_()`       | Full daily notification                                              |
| `sendSyncNotification_(ctx, ids)`   | Immediate notification for new appointments                          |
| `updateConsecutivePaidCounts_(ctx)` | Recalculates `Consecutive Paid` streaks for all clients              |
| `updateNoShowLateCounts_(ctx)`      | Recalculates `No Show` and `Late` counts (12m) for all clients       |
| `setupTriggers()`                   | Installs onEdit + 5-min sync triggers in one click                   |
| `setupOnOpenSync()`                 | Installs installable onOpen trigger for sync on sheet open           |
| `validateSetup()`                   | Checks calendar, sheets, Telegram credentials, and triggers          |
| `nameCase_(v)`                      | Normalises names to Title Case for comparison                        |
| `toSheetDateTime_(s, tz, isAllDay)` | Converts Calendar dates to Sheet cell format                         |

## Script Properties

| Key                   | Purpose                                                          |
| --------------------- | ---------------------------------------------------------------- |
| `TELEGRAM_BOT_TOKEN`  | KashmirBarberBot token                                           |
| `TELEGRAM_CHAT_ID`    | Telegram chat id                                                    |
| `CALENDAR_SYNC_TOKEN` | Auto-managed incremental sync state; delete to force full resync |

## Useful Manual Functions

Run these from the Apps Script editor dropdown:

- `syncThisYear` — syncs entire current calendar year
- `setupNotificationTrigger` / `removeNotificationTrigger` — manage 9 PM daily trigger
- `setupIncrementalSync` / `removeIncrementalSync` — manage 5-min trigger
- `runMigration` — one-time migration to populate the `Cached Name` column (L=12)
- `cleanupDuplicates` — removes duplicate rows safely
- `testTelegramAPI` — tests Telegram bot connection
- `debugTomorrow` — logs tomorrow's appointments to console
- `formatSpreadsheet` — applies dark/charcoal visual theme with pastel accents; safe to re-run

## Debugging in Apps Script

**Never use `safeAlert_()` or `SpreadsheetApp.getUi().alert()` for debug output.** When a function is run from the Apps Script editor (not from the sheet menu), there is no UI context — `alert()` blocks indefinitely waiting for a click that never comes, making it look like a timeout.

When adding debug logging to any function:
- Use `Logger.log(...)` for step-by-step output — visible in **View → Logs** after the run
- Use `SpreadsheetApp.getActive().toast(message, title, durationSeconds)` if you need visible output — works from both the editor and the sheet menu, and auto-dismisses
- Never replace or wrap `safeAlert_` calls with another blocking UI call

## Personal Preferences

- I like the barbersheet to be as minimalistic as possible
- The barbersheet menu should not be filled with actions/functions
