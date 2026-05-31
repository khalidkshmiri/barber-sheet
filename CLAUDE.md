# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What This Is

A Google Apps Script for a barber business management system. The single script file (`barber-sheet-script.js`) runs inside a Google Sheets spreadsheet and syncs appointments from Google Calendar (which syncs from Apple Calendar via iCloud), manages client records, and sends Telegram notifications.

**Owner:** Khalid — barber based in Rotterdam, Netherlands

## Deployment

There are no build or test commands. Development workflow:
1. Edit `barber-sheet-script.js` locally
2. Paste into the Apps Script editor: **Extensions → Apps Script** in the Google Sheet
3. To deploy as web app: **Deploy → Manage deployments → Edit → New version → Deploy**
4. To run functions manually: select from the function dropdown in the editor and click Run

Syntax checking: `node --check barber-sheet-script.js`

## Google Sheet Structure

Four sheets: **Appointments**, **Clients**, **Services**, **Subscriptions**

**Appointments columns (1-based):**
A=Date, B=Time, C=Name (XLOOKUP formula), D=Price, E=Payment, F=Status, G=Tips, H=Late (checkbox), I=Notes, J=Service, K=ClientID, L=EventID (iCalUID), M=CachedName (fallback for XLOOKUP)

**Clients columns (1-based):**
A=Name, B=FavService, C=LastVisit, D=SocialMedia, E=Notes, F=NoShow(12m), G=Late(12m), H=Referral, I=TotalVisits, J=TotalTips, K=TotalSpent, L=ClientID, M=FirstVisit, N=DoNotCut (checkbox), O=ConsecutivePaid (auto), P=VIP (checkbox)

**Locale:** Dutch — Google Sheets formulas use **semicolons** as separators, not commas. Example: `=XLOOKUP(K2; Clients!L:L; Clients!A:A; M2)`

## Architecture

### Calendar Sync

Calendar event title format: `ClientName - Service` (e.g. `Ibrahim - Haircut`)

**iCalUID vs opaque ID (critical):** `CalendarApp.event.getId()` returns iCalUID (`abc123@google.com`). The Calendar REST API's `item.id` returns an opaque ID (`abc123`). These differ for the same event. The sheet always stores iCalUID in col L. Incremental sync must use `item.iCalUID` (not `item.id`) — this was a bug that caused duplicate rows; it is fixed.

**Three sync paths:**
- `syncCalendarToSheets()` — manual full sync, ±14 days window, ~18 sec
- `syncCalendarIncremental_()` — 5-min trigger using Calendar API sync token; exits in ~0.4 sec if nothing changed; uses `LockService` to prevent concurrent runs
- `doGet(e)` — web app endpoint called by iOS Shortcut when Calendar app closes; calls `syncCalendarIncremental_()` and shares its sync token

After any sync: `updateUpcomingToNotPaid_()`, `updateConsecutivePaidCounts_()`, `updateNoShowLateCounts_()`, `sortAndHideAppointments_()` always run.

### Subscription Logic

When service is "Monthly Subscription": sets `serviceToWrite = "Haircut"`, price = 0, payment = "Subscription", and calls `createSubscriptionEntry_()`. A client with active subscription credits also gets price = 0, payment = "Subscription" automatically. Credits formula in Subscriptions sheet counts used Subscription-paid appointments since start date (max 4 per month).

### onEdit Trigger (`processSheetChanges`)

This is an **installable trigger** (not a simple `onEdit`). It handles:
- Dashboard C2 checkbox → triggers `syncCalendarIncremental_()` + sends a Telegram confirmation
- Appointments col E (Payment) change → auto-sets Status and Price. Valid payment values: Cash, Tikkie, Subscription, Free
- Appointments col F (Status) change → No Show/Cancelled clears Late checkbox; Free-* sets price to €0
- Clients col A → auto-formats name to Title Case

**Payment = "Free"** sets price → 0, status → "Paid", and counts toward the consecutive paid streak.

### Telegram Notifications

Script Properties store `TELEGRAM_BOT_TOKEN` and `TELEGRAM_CHAT_ID`. Two notification functions:
- `sendDailyNotification()` — called by 9 PM daily trigger; shows tomorrow's appointments, today's unpaid, unreliable client alerts
- `sendSyncNotification_(ctx, newEventIds)` — fires immediately when new appointments are added during sync

Reliability badges based on combined NoShow + Late count from Clients sheet: ⛔ DO NOT CUT (col N checkbox), ⚠️ Unreliable (3+), 🟡 Watch (1–2), ✅ Reliable (0).

Consecutive paid streak (col O) resets on No Show or Late checkbox; at 5+ shows "✅ ELIGIBLE FOR FREE" in notifications.

## Key Functions Reference

| Function | Purpose |
|----------|---------|
| `upsertEvents_(events, ctx)` | Core sync logic — processes calendar events into sheet rows |
| `prepareContext_()` | Loads all sheet references and indexes into a context object |
| `syncCalendarIncremental_()` | Incremental sync using Calendar API sync token |
| `processSheetChanges(e)` | Installable onEdit handler |
| `sendTelegramNotification_()` | Full daily notification |
| `sendSyncNotification_(ctx, ids)` | Immediate notification for new appointments |
| `updateConsecutivePaidCounts_(ctx)` | Recalculates col O streaks for all clients |
| `updateNoShowLateCounts_(ctx)` | Recalculates col F (NoShow 12m) and col G (Late 12m) for all clients |
| `setupTriggers()` | Installs onEdit + 5-min sync triggers in one click |
| `setupOnOpenSync()` | Installs installable onOpen trigger for sync on sheet open |
| `validateSetup()` | Checks calendar, sheets, Telegram credentials, and triggers |
| `nameCase_(v)` | Normalises names to Title Case for comparison |
| `toSheetDateTime_(s, tz, isAllDay)` | Converts Calendar dates to Sheet cell format |

## Script Properties

| Key | Purpose |
|-----|---------|
| `TELEGRAM_BOT_TOKEN` | KashmirBarberBot token |
| `TELEGRAM_CHAT_ID` | `8005583266` |
| `CALENDAR_SYNC_TOKEN` | Auto-managed incremental sync state; delete to force full resync |

## Useful Manual Functions

Run these from the Apps Script editor dropdown:
- `syncThisYear` — syncs entire current calendar year
- `setupTriggers` — installs onEdit + 5-min sync triggers (run once after initial paste)
- `setupOnOpenSync` — syncs on sheet open (installable onOpen trigger)
- `setupNotificationTrigger` / `removeNotificationTrigger` — manage 9 PM daily trigger
- `setupIncrementalSync` / `removeIncrementalSync` — manage 5-min trigger
- `validateSetup` — checks calendar, sheets, Telegram, triggers
- `runMigration` — one-time migration to populate col M (CachedName)
- `cleanupDuplicates` — removes duplicate rows safely
- `testTelegramAPI` — tests Telegram bot connection
- `debugTomorrow` — logs tomorrow's appointments to console

## Personal Preferences
- I like the barbersheet to be as minimalistic as possible
- The 