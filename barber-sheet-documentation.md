# Barber Sheet System — Full Documentation

This document contains everything needed to continue development in a new chat. Hand this to Claude and say: "Read this document and continue where we left off."

---

## Overview

A Google Sheets system that manages a barber business. Appointments sync from Apple Calendar (via Google Calendar) using a Google Apps Script. Telegram notifications provide client intelligence before each appointment.

**Owner:** Khalid — barber based in Rotterdam, Netherlands
**Goal:** Eventually own a barber shop

---

## Current File

**`barber_sheet_script.js`** — this is the one and only script file. Always work from this.

---

## Google Sheet Structure

### Sheet: Appointments
| Col | Letter | Content |
|-----|--------|---------|
| A | Date | Appointment date |
| B | Time | Appointment time |
| C | Name | `=XLOOKUP(K{row}; Clients!L:L; Clients!A:A; M{row})` formula |
| D | Price | € amount |
| E | Payment | Cash / Tikkie / Subscription / (blank) |
| F | Status | Paid / Not Paid / Upcoming / No Show / Cancelled / Free - * |
| G | Tips | Tip amount |
| H | Late | Checkbox |
| I | Notes | Free text notes |
| J | Service | Haircut / Monthly Subscription / etc. |
| K | ClientID | Unique client ID |
| L | EventID | Google Calendar iCalUID (e.g. `abc123@google.com`) |
| M | CachedName | Fallback name if XLOOKUP fails |

### Sheet: Clients
| Col | Letter | Content |
|-----|--------|---------|
| A | Name | Client name |
| B | FavService | Favourite service |
| C | LastVisit | Auto-updated |
| D | SocialMedia | Social media handle |
| E | Notes | Client notes |
| F | NoShow | No-show count (12 months) |
| G | Late | Late count (12 months) |
| H | Referral | Who referred them |
| I | TotalVisits | Total visit count |
| J | TotalTips | Total tips received |
| K | TotalSpent | Total amount spent |
| L | ClientID | Unique ID (used for XLOOKUP) |
| M | FirstVisit | Date of first visit |
| N | DoNotCut | Checkbox — marks client as banned |
| O | ConsecutivePaid | Auto-calculated streak count |
| P | VIP | Checkbox — manual VIP flag |

### Sheet: Subscriptions
| Col | Content |
|-----|---------|
| Name, Price, PaymentType, StartDate, Credits (formula), Status, ExpiryDate, Notes, ClientID |

### Sheet: Services
- Haircut = €15
- Monthly Subscription = €40
- Free = €0

**Important:** Dutch locale — Google Sheets formulas use semicolons as separators, not commas.

---

## Calendar → Sheet Sync

**Calendar:** "Barber Appointments" in Google Calendar (synced from Apple Calendar via iCloud)

**Event title format:** `ClientName - Service`
Examples: `Ibrahim - Haircut`, `Rajko - Monthly Subscription`

**How it works:** Script reads calendar events and creates/updates rows in Appointments sheet. EventID (iCalUID) stored in col L is used to match existing rows and prevent duplicates.

**Critical — iCalUID vs opaque ID:**
- `CalendarApp.event.getId()` returns iCalUID: `abc123@google.com`
- Calendar REST API `item.id` returns opaque ID: `abc123`
- These are different strings for the same event
- The sheet stores the iCalUID format
- Incremental sync must use `item.iCalUID` not `item.id` (bug was fixed)

---

## Sync Methods

### Sync Now (manual)
- Function: `syncCalendarToSheets()`
- Fetches all events from -30 days to +90 days
- Runs: consecutive paid count update, sorting, hiding
- Sends Telegram notification if new/cancelled appointments
- Time: ~18 seconds
- Quota: ~18 sec per run

### Incremental Sync (auto, 5 min)
- Function: `syncCalendarIncremental_()`
- Uses Google Calendar API sync token — only fetches what changed
- If nothing changed: exits in ~0.4 sec
- If something changed: processes only those events (~3–5 sec)
- Sends Telegram notification if new appointments added
- Quota: ~1.9 min/day (very safe, Google free limit is 6 min/day)
- Uses `LockService` to prevent concurrent runs (web app + trigger firing simultaneously)
- Token expires after ~7 days → auto falls back to full sync + gets new token

**Setup (one time):**
1. Apps Script editor → Services (+) → Google Calendar API → Add
2. Reload sheet → Barber Tools → Setup Incremental Sync (5 min)

### Web App (iOS Shortcut)
- Function: `doGet(e)`
- Calls `syncCalendarIncremental_()` — shares sync token with 5-min trigger
- Returns plain text: `✅ Sync complete!`
- Used by iOS Shortcut automation: closes Calendar app → silently calls web app URL

**iOS Shortcut setup:**
- Shortcuts app → Create shortcut → Add "Get Contents of URL" action
- Paste web app URL, Method: GET
- Set as automation: when Calendar app Is Closed
- "Ask Before Running" OFF → runs silently in background

**Deploying web app:**
- Apps Script → Deploy → Manage deployments → Edit → New version → Deploy
- URL stays the same after each deployment

---

## Barber Tools Menu (in Google Sheet)

| Item | Function |
|------|----------|
| 🔄 Sync Now | `syncCalendarToSheets` |
| ⚡ Setup Incremental Sync (5 min) | `setupIncrementalSync` |
| 🗑️ Stop Incremental Sync | `removeIncrementalSync` |

**All other functions can be run directly from Apps Script editor:**
Extensions → Apps Script → select function from dropdown → Run

Useful functions to run manually:
- `syncThisYear` — syncs entire calendar year
- `setupNotificationTrigger` — sets up daily 9 PM Telegram notification
- `removeNotificationTrigger` — stops daily notification
- `sendDailyNotification` — sends notification immediately (for testing)
- `runMigration` — one-time migration to populate col M with cached names
- `testTelegramAPI` — tests Telegram connection

---

## Telegram Notifications

**Bot:** KashmirBarberBot
**Stored in Script Properties:**
- `TELEGRAM_BOT_TOKEN` — bot token
- `TELEGRAM_CHAT_ID` — `8005583266`

### Daily Notification (9 PM)
Function: `sendTelegramNotification_()`
Three sections (only sent if at least one section has content):

**1. Tomorrow's appointments**
- Time, name, price, reliability badge, VIP ⭐
- Previous appointment info: "Last: Paid (€15) — notes"
- DoNotCut ⛔, loyalty tier, free haircut eligibility

**2. Today's unpaid**
- List of today's appointments with no payment recorded

**3. Unreliable alerts**
- Clients with 3+ NoShows/Lates appearing in 1–2 days
- Deduplicated — not shown if already in tomorrow's section

### Sync Notification (on new appointment)
Function: `sendSyncNotification_(ctx, newEventIds)`
Triggered when: incremental sync or manual sync detects newly added appointments.

Shows for each new appointment (regardless of how far in the future):
- Date and time
- Client name + VIP badge + reliability badge + loyalty label
- Price
- Last appointment status + notes
- "First visit" if no previous appointments

**Why:** Khalid needs to decide whether to keep a booking based on client reliability — so the notification fires immediately on sync, not just at 9 PM.

### Reliability Badges
- ⛔ DO NOT CUT — col N checkbox
- ⚠️ Unreliable (3+ NoShow + Late combined)
- 🟡 Watch (1–2 NoShow + Late combined)
- ✅ Reliable (0)

---

## Status Values

| Status | Meaning | Behaviour |
|--------|---------|-----------|
| Upcoming | Future appointment, not yet processed | Auto-set for future events |
| Not Paid | Past appointment, payment not received | Auto-set when date passes |
| Paid | Payment received | Set manually or via Payment column trigger |
| No Show | Client didn't show | Clears Late checkbox |
| Cancelled | Appointment cancelled | Clears Late checkbox |
| Free - Family | Free cut (family) | Sets price to €0 |
| Free - Charity | Free cut (charity) | Sets price to €0 |
| Free - Friend | Free cut (friend) | Sets price to €0 |
| Free - Loyalty Reward | Free cut (loyalty) | Sets price to €0 |
| Free - Other | Free cut (other) | Sets price to €0 |

---

## Good Client Rewards

### VIP Badge (manual)
- Clients sheet col P checkbox
- Shows ⭐ in Telegram notification
- Does not affect formulas or prices

### Loyalty Tier (auto — Consecutive Paid)
- Clients sheet col O "Consecutive Paid" — auto-updated on every sync
- Counts consecutive paid appointments from most recent backwards
- **Resets to 0** if most recent appointment is No Show OR has Late checkbox checked
- When count ≥ 5: shows "✅ ELIGIBLE FOR FREE" in notification
- Free cuts (Free - *) count toward the streak
- No Show and Late reset it

---

## processSheetChanges (installable onEdit trigger)

Fires when certain cells are edited:

| Trigger | Action |
|---------|--------|
| Dashboard C2 checkbox | Triggers sync |
| Appointments col E (Payment) changes | Auto-sets Status to Paid/Upcoming/Not Paid; auto-sets Price for Subscription |
| Appointments col F (Status) changes | No Show/Cancelled clears Late; Free-* sets Price to €0, keeps Late |
| Clients col A | Auto-formats to Title Case |

---

## Key Helper Functions (internal, prefixed with `_`)

- `parseEventTitle_(title)` — splits "Name - Service" into parts
- `loadClientsIndex_(sheet)` — Map of name → row for fast lookup
- `loadAppointmentEventIdIndex_(sheet)` — Map of eventId → row number
- `loadActiveSubscriptionsIndex_(sheet)` — finds clients with active subscription credits
- `loadClientNotificationMap_(sheet)` — loads client data for notification building
- `getClientLastAppointment_(name, sheet)` — returns most recent past appointment
- `getReliabilityInfo_(name, clientMap)` — returns badge and colour
- `isVIP_(name, sheet)` — checks col P checkbox
- `getConsecutivePaidCount_(name, sheet)` — reads col O
- `updateConsecutivePaidCounts_(ctx)` — recalculates all streaks
- `updateUpcomingToNotPaid_(sheet)` — flips past "Upcoming" rows to "Not Paid"
- `sortAndHideAppointments_(sheet)` — sorts by date desc, hides old rows
- `upsertEvents_(events, ctx)` — processes calendar events into sheet rows; returns `{ newCount, updatedCount, cancelledCount, newEventIds }`
- `sendSyncNotification_(ctx, newEventIds)` — immediate notification for new appointments
- `sendTelegramNotification_()` — full daily notification (tomorrow / unpaid today / unreliable)
- `safeAlert_(msg)` — shows UI alert safely (won't crash if no UI context)
- `startOfDay_(date)`, `endOfDay_(date)`, `addDays_(date, n)` — date utilities
- `toSheetDateTime_(date, tz, isAllDay)` — converts calendar date to sheet format
- `nameCase_(name)` — normalises name to lowercase for comparison

---

## Constants (top of script)

```javascript
const APPOINTMENTS_SHEET  = "Appointments";
const CLIENTS_SHEET       = "Clients";
const SUBSCRIPTIONS_SHEET = "Subscriptions";
const SERVICES_SHEET      = "Services";
const DAYS_BACK           = 30;
const DAYS_FORWARD        = 90;
```

---

## Script Properties (stored in Apps Script)

| Key | Value |
|-----|-------|
| `TELEGRAM_BOT_TOKEN` | Bot token for KashmirBarberBot |
| `TELEGRAM_CHAT_ID` | `8005583266` |
| `CALENDAR_SYNC_TOKEN` | Stored after first incremental sync run; auto-managed |

---

## Known Issues / History

### Duplicate appointments bug (FIXED)
**Cause:** Incremental sync used `item.id` (opaque REST API ID) to match events, but the sheet stores `item.iCalUID` (set by `CalendarApp.event.getId()`). These are different strings for the same event, so every event looked new and got a duplicate row.

**Fix applied:** Incremental sync now uses `item.iCalUID` with fallback to `item.id`. The API call explicitly requests `iCalUID` in the fields parameter.

### Concurrent sync runs (FIXED)
**Cause:** iOS Shortcut and 5-min trigger could fire at the same time, both processing the same events.

**Fix applied:** `LockService.getScriptLock()` at the start of `syncCalendarIncremental_()`. If already running, second call exits immediately. Also, `doGet()` now calls `syncCalendarIncremental_()` instead of the full sync, so they share the same sync token.

### Syntax error in sendSyncNotification_ (FIXED)
**Cause:** Template literals with emoji and newlines were written as Python strings with embedded `\n`, which became actual newlines inside JS string literals when written to file.

**Fix applied:** Function rewritten using `lines.join("\n")` pattern with unicode escapes for emoji. Node `--check` verified zero syntax errors.

---

## Quota Analysis

| Method | Time/run | Runs/day | Total/day | Safe? |
|--------|----------|----------|-----------|-------|
| Full sync (manual) | ~18 sec | ~3 | ~54 sec | ✅ |
| Old auto-sync (15 min) | ~4 sec | 96 | ~6.4 min | ⚠️ borderline |
| Incremental (5 min, nothing changed) | ~0.4 sec | 288 | ~1.9 min | ✅ |
| Incremental (5 min, something changed) | ~3–5 sec | rare | minimal | ✅ |

Google free account trigger quota: 6 min/day

---

## Dashboard Formulas (Dutch locale — semicolons)

```
This week start:  =TODAY()-MOD(TODAY()-2;7)
This week end:    =TODAY()-MOD(TODAY()-2;7)+6
Paid count:       =COUNTIFS(Appointments!A:A;">="&StartDate;Appointments!A:A;"<="&EndDate;Appointments!F:F;"Paid")
```

---

## What Still Works / Pending Items

- ✅ Manual sync (Sync Now button)
- ✅ Incremental sync (5 min, iCalUID fix applied)
- ✅ iOS Shortcut automation (Calendar closes → syncs silently)
- ✅ Daily 9 PM Telegram notification (3 sections)
- ✅ Sync notification for new appointments (any date, full client details)
- ✅ Reliability badges (⛔ / ⚠️ / 🟡 / ✅)
- ✅ VIP badge (⭐, manual)
- ✅ Consecutive paid streak + free haircut eligibility
- ✅ Free haircut types (Family / Charity / Friend / Loyalty Reward / Other)
- ✅ Last appointment info in notifications
- ✅ Subscription handling
- ✅ processSheetChanges onEdit trigger
- ⬜ Conditional formatting for Status column (discussed but not confirmed implemented): Not Paid = Orange, Paid = Green, No Show = Red, Cancelled = Gray, Upcoming = Blue
