---
name: mc-repo-guide
description: Complete working knowledge of the Meaningful Conversations St. Louis Google Sheets + Apps Script automation - the Contact List sheet layout (rows, columns, section markers, RSVP strings), every constant and helper, the multi-account web-app pattern, triggers, deployment via clasp, and the conventions new scripts must follow. Use whenever reading, writing, or debugging scripts for this spreadsheet, or when rebuilding the project from scratch.
---

# Meaningful Conversations — Repo & Spreadsheet Guide

This skill contains everything needed to work on (or rebuild from scratch) the
Apps Script automation for the **Meaningful Conversations St. Louis** event and
contact-management spreadsheet. All code lives in `scripts/*.gs` +
`scripts/*.html` in one Apps Script project bound to the spreadsheet.

## 1. The spreadsheet — sheets (tabs)

| Sheet name | Constant | Purpose |
|---|---|---|
| `Contact List` | `SHEET_NAMES.CONTACT_LIST` | The master sheet — people × events grid. Everything keys off this. |
| `Schedule` | `SHEET_NAMES.SCHEDULE` | Source of upcoming events. Col C = date (often string like "Mon, 01/12/26 6:00 PM"), col D = title, col E = location/address. Skip-words: `OFF`, `TBD`, `TOPIC` (`SCHEDULE_SHEET_CONSTANTS.SKIP_WORDS`). |
| `EventBrite Import` | `SHEET_NAMES.EVENTBRITE` | Staging area. Rows A..L (12 cols, `HELPER_CONSTANTS.EVENTBRITE_COLUMN_COUNT`), data from row 2. The move script sweeps ALL rows into the Contact List and deletes them. Facebook CSV imports stage here too (Platform "FB"). |
| `STL MO Barcode (signup sheet)` | `SHEET_NAMES.SIGNUP` | Barcode/form sign-ins at the door. Cols (0-based, `SIGNUP_COLS`): A timestamp, B name, C email, D phone, E comments. Feeds attendance marking. |
| `Email Tracking` | `BULK_EMAILER_DEFAULTS.TRACKING_SHEET_NAME` | Bulk Emailer log. Header: Email, Sent Status, Run Type, Name, Timestamp, Error, Attachment, Subject. Idempotency key = email + subject (col H). Batch rows stack many addresses newline-separated in the Email cell. |
| `Lifecycle Email Tracking` | config `TRACKING_SHEET_NAME` | Lifecycle emailer log. Idempotency key = template + email + eventKey (title\|date). |
| `Event IDs` | (literal) | B = event title, C = event code. Row-7 header formulas VLOOKUP into it; `lookupEventCode_` reads it. |

## 2. Contact List layout — the critical map

### Header rows (1-based, `ROW_NUMBERS.ROW_1..ROW_12`)
- **Row 4**: room info (e.g. "Rm 208") — used in email bodies.
- **Row 5**: column headers — "very important row". Counter/landmark columns are found by scanning this row for `COL_CONSTANTS` labels: `# Events Attended`, `# Events RSVP'd`, `Next Steps`, `Total RSVP'd`, `Total Attended`, `Total Attended w/o RSVP`.
- **Row 6**: event **dates** (Date values) for each event column.
- **Row 7**: event **titles** (usually `=VLOOKUP(X7,'Event IDs'!$B:$C,2,FALSE)` via `FORMULAS.V_LOOKUP`).
- **Row 12** (`ROW_NUMBERS.ROW_12`): first data row for bulk scans. Sorting/sections treat **row 13** as the first sortable data row.

### Identity columns A–N (0-based `COLUMN_INDEX` in helpers.gs)
| Idx | Col | Name | Notes |
|---|---|---|---|
| 0 | A | `FULL_NAME_KEY` | `=IF(ISTEXT(C12),CONCATENATE(C12," ",D12),"")` (`FORMULAS.CONCATENATE_NAME`). **Merge key.** |
| 1 | B | `NAME_DISPLAY` | Also holds event-block titles and section markers (see below). |
| 2 | C | `FIRST_NAME` | |
| 3 | D | `LAST_NAME` | |
| 4 | E | `STATUS` | e.g. "repeat attendee" (`BULK_EMAILER_DEFAULTS.REPEAT_FLAG_VALUE`). |
| 5 | F | `EMAIL` | May hold MULTIPLE emails, comma/semicolon-separated. Always split on `/[,;]+/` and validate each. |
| 6 | G | `PHONE` | Dash format `314-123-4567` (`formatPhoneWithDashes_`). Extra numbers append comma-separated. |
| 7 | H | `NEIGHBORHOOD` | |
| 8 | I | `AGE` | approximate |
| 9 | J | `SIGNUP_DATE_TIME` | Original signup date/time. **Never merge/concat.** |
| 10 | K | `SIGNUP_PLATFORM` | "EB", "FB", "MU", "Walk-in". **Never merge/concat.** |
| 11 | L | `SIGNUP_EVENT_TITLE` | **Never merge/concat.** Titles can contain commas ("One God, Many Paths") — never naively split on comma; match against Row-7 titles. |
| 12 | M | `SIGNUP_EVENT_CODE` | **Never merge/concat.** |
| 13 | N | `DESCRIPTION` | |

### Event columns (O onward)
- Start at column **O = 15** (`HELPER_CONSTANTS.EVENT_NAMES_START_COL`). Newest events are inserted at the LEFT (a new column is inserted after O, becoming P… the code inserts `insertColumnAfter(15)`).
- Each cell holds one of the **exact** `RSVP_DROP_DOWN_CONSTANTS` strings (note the literal `\n`):
  - `"-"`, `"--"` (no involvement)
  - `"rsvp'd: yes\nattended: ?"`, `"rsvp'd: maybe\nattended: ?"`, `"rsvp'd: no\nattended: ?"`
  - `"rsvp'd: yes\nattended: yes"`, `"rsvp'd: no\nattended: yes"`, `"rsvp'd: maybe\nattended: yes"`, `"rsvp'd: -\nattended: yes"`
  - `"rsvp'd: yes\nattended: no"`, `"rsvp'd: no\nattended: no"`
- Cells get a dropdown validation built from `Object.values(RSVP_DROP_DOWN_CONSTANTS)`.
- Counting formulas match substrings: attended = `COUNTIF(...,"*Attended: yes")`, rsvp = `COUNTIF(...,"RSVP'd: yes*")` (see `FORMULAS`).
- To the right of the last event column sit the counter columns found via Row 5: `# Events Attended`, `# Events RSVP'd`, `Next Steps`.

### Section markers & vertical layout (top to bottom)
1. **Attended section**: rows **13** up to the row above the `"RSVP 2+"` marker (col B). Sorted by `sortAttendedRows`.
2. **`RSVP 2+`** marker row (col B, `COL_CONSTANTS.RSVP_2_PLUS`).
3. **RSVP section**: from `RSVP2+ row + 2` down to `Stop RSVP row − 2`. Sorted by `sortRSVPRows`.
4. **`Stop RSVP`** marker row (col **A**, `COL_CONSTANTS.STOP_RSVP`).
5. **Main list**: event blocks. Each block = title in col B (matching a Row-7 title after paren-stripping), attendee rows starting **two rows below the title**, then total rows (`Total RSVP'd` label in col B) and a collapsible row group. `"ATTENDED 1+"` (col B) marks the attended-section header near the top.
- Sections at the top are **copies** of main-list rows — dedup happens by email, and the merge script collapses duplicates by the col-A key.
- The old `"Start Email"` / `"Stop Email"` col-A markers are **retired** — audience selection is by section now.

### DeveloperMetadata markers (fast lookups)
`MARKER_KEYS` pins landmarks with DeveloperMetadata for O(1) lookup that survives row/col inserts: `mc.col.eventsAttended`, `mc.col.eventsRsvpd`, `mc.col.nextSteps` (row-5 headers), `mc.row.totalRsvpd` (current block's total row). Lookups are **self-healing**: `findColMarker_(sheet, key, label)` / `findRowMarker_` verify the pin still sits on the label, else rescan and re-pin. Column pins must use whole-column A1 ranges (`"P:P"`). Always locate counter columns this way — never hardcode.

## 3. File-by-file map (`scripts/`)

| File | What it does |
|---|---|
| `helpers.gs` | ALL shared constants (`SHEET_NAMES`, `COL_CONSTANTS`, `COLUMN_INDEX`, `RSVP_DROP_DOWN_CONSTANTS`, `HELPER_CONSTANTS`, `ROW_NUMBERS`, `FORMULAS`, `UI_CONSTANTS`, `SCHEDULE_SHEET_CONSTANTS`, `EMAILER_KEYS`, `MARKER_KEYS`), `columnToLetter`, `normalizeString` (apostrophes), `normalizeByStrippingWhiteSpaceAtTheEnd` (trim+collapse+lowercase), `sheetsByName()` → `[contactListSheet, eventbriteSheet, scheduleSheet]`, marker finders/pinners. |
| `menu_items.gs` | `onOpen()` builds two menus: "Custom Actions" (sheet maintenance) and "Email" (Email Composer, Bulk Emailer, No-Email Report). |
| `nightly_trigger.gs` | `nightlyTrigger()`: importNewEventsFromSchedule → EventBrite move → merge → copy attended/RSVP sections → sorts → markNoShows → calendar sync. |
| `insert_new_col_and_row.gs` | `importNewEventsFromSchedule()`: finds first Schedule event not in Row 7 (title match + date within 7 days), inserts a new event column after O (formatting copied from old P, date = previous + 7 days skipping OFF weeks) and a new event block (title row + totals row + group) in the main list. Idempotent — no-ops when nothing new. |
| `move_from_eventbrite_to_contact.gs` | Batch-moves staged rows into event blocks: bottom-up, inserts **two rows below the block title**, expands row groups first, writes values (B..L → C..M), name/count formulas, RSVP dropdown = `rsvp'd: yes\nattended: ?`, dashes to the right, formatting, deletes staged rows bottom-up. |
| `mark_attendance_from_signup.gs` | Barcode sheet → attendance. `onSignupFormSubmit(e)` (form-submit trigger) and `markAttendanceFromSignupSheet()` (manual batch). Match by email → phone → fuzzy name (exact/spaceless/first+last swapped). Flips `attended: ?/no` → `yes`; empty cell → `rsvp'd: no\nattended: yes`. No match at all → inserts a **walk-in row** into today's event block (Platform "Walk-in"). Updates phone if new. `setupSignupFormTrigger()` creates the trigger. |
| `merge_duplicates.gs` | `mergeRowsByKeyPreserveAllFormulas()`: merges rows sharing the col-A key when F/G are compatible (equal-or-empty, case-insensitive). Skip-set (top row wins, never concatenated): A, B, C, D, **J, K, L, M**, counter cols. Event cells merge by RSVP priority (`pickByRsvpPriority_` — attended:yes > attended:no > `?` > dash), never concatenated. Other cells concat with ", " only when genuinely different. Deletes dupes bottom-up in consecutive-run batches. |
| `move_to_attended_section.gs` / `move_to_rsvp_section.gs` | Copy rows with #attended ≥ 1 / #rsvp ≥ 2 into their sections (formulas re-based via `adjustFormulaForRow`), then merge. |
| `sort_attended.gs` / `sort_rsvp.gs` | Sort the two sections by event columns (O.. → Next Steps col), ascending. |
| `mark_no_shows.gs` | After events pass, flips stale `attended: ?` to `attended: no`. |
| `backfill_events_formulas.gs` | Rewrites count formulas when columns move. |
| `calendar_sync.gs` | Syncs Schedule events to Google Calendar (6:30–8:00 PM window). |
| `emailer.gs` | **Bulk Emailer** — see §5. |
| `lifecycle_emailer.gs` | **Lifecycle emails** — see §6. |
| `email_composer.gs` + `email_composer_dialog.html` + `email_composer_send_page.html` | Email Composer UI + web-app router (`doGet`) — see §6/§7. |
| `bulk_emailer_dialog.html` | Bulk Emailer UI (launcher popup + full-tab form). |
| `facebook_csv_import.gs` + `facebook_csv_import_dialog.html` | Facebook guest CSV → EventBrite-shaped staging rows ("Going" only) → normal move. `doPost` handler (`action=fbimport`). |
| `cleanup_signup_origin.gs` | One-shot repair of conjoined K/L/M values (keeps left-most; col L matched against Row-7 titles because titles contain commas). |
| `no_email_report.gs` | **No-Email Report** — groups an audience's rows by person and lists everyone with no valid address in col F, with phone / signup platform (K) / signup event (L) to reach them another way. Rebuilds the "No Email Report" sheet each run; says "Zero" explicitly when everyone is reachable. Menu item `showNoEmailReport` (whole list) + `runNoEmailReport(audience)` from the Bulk Emailer. |
| `charts.gs` | RSVP vs attendance charts. |
| `appsscript.json` | Timezone America/Chicago, V8, webapp `executeAs: USER_ACCESSING`, `access: ANYONE`, Calendar advanced service, scopes incl. `script.send_mail`, `script.external_request`. |

## 4. The multi-account web-app pattern (CRITICAL)

**Problem**: in-sheet popup dialogs run in an iframe whose `google.script.run`
binds to the browser's DEFAULT Google session → `PERMISSION_DENIED` when the
user is signed into several accounts.

**Solution**: a web-app deployment pinned to the team account.
- `COMPOSER_WEBAPP_URL` (in `email_composer.gs`) = the stable `/exec` URL.
- `COMPOSER_SEND_ACCOUNT` = `st.louis@meaningful-conversations.org`.
- Popups either (a) hop actions through a URL/form-POST opened with
  `&authuser=<team account>` in a new tab, or (b) are mere **launchers** that
  open the full UI as a web-app tab (`?page=...&authuser=...`) where
  `google.script.run` IS session-safe.
- `doGet` routes: no params → Email Composer tab; `action=send` → composer
  sending page; `page=facebook` → FB import tab; `page=bulkemailer` → Bulk
  Emailer tab. `doPost` routes: `action=fbimport` (big payloads must POST).
- Dialog HTML templates get data **injected at render time**
  (`t.bootData = JSON.stringify(data).replace(/</g, "\\u003c")`, read with
  `JSON.parse(<?!= JSON.stringify(bootData) ?>)`) so loading never needs
  `google.script.run`. Each template gets `uiMode` = `"dialog"` or `"webapp"`.
- **After changing any web-app-served file you MUST redeploy the pinned
  deployment** (see §8) — `clasp push` alone only updates HEAD.

### Live progress pattern
Long sends write snapshots to `CacheService.getScriptCache()` under key
`"composerProgress:" + token` (TTL 600s), shape:
`{ total, processed, sent, skipped, failed, planned, done }`.
Writers: `lifecycleProgressUpdate_` (lifecycle) and `bulkProgress*_` (bulk).
Reader: `getComposerSendProgress(token)`; UIs poll it every ~1.5s while the
send runs via `google.script.run`, showing "Processing X of Y". Tokens are
client-generated random strings (or `Utilities.getUuid()`).

## 5. Bulk Emailer (`emailer.gs`)

- Entry: `sendBulkEmails(payload)` from the UI; menu launcher `showBulkEmailerDialog()`.
- **Audience** (`bulkAudienceBounds_`, 0-based inclusive bounds mirroring the sorts): `whole` = rows 13..last; `attended` = 13..(RSVP2+ row − 1); `rsvp` = (RSVP2+ row + 2)..(StopRSVP row − 2).
- Body is TYPED (no Google Doc): `buildBulkEmailHtml_` escapes it and wraps in a serif card (blank line → paragraph, newline → `<br>`), footer "Meaningful Conversations · St. Louis, MO".
- Attachment uploaded from browser as `{name, mimeType, dataB64}` → `Utilities.newBlob(Utilities.base64Decode(...))`. 20 MB client-side guard.
- Filters: repeat-attendee (col E), per-event (Row-7 title → column, non-dash cell), attended-count strictly-more/less-than, exclude list, and **active on/after an event** (`sinceEventKey` = `title|yyyy-MM-dd` → `sinceEventColumns_` resolves that event and every later one by Row-6 date; `sinceMode` `"attended"` | `"any"` = attended-yes vs. also RSVP yes/maybe). That's the "email everyone since <event>" audience — no row numbers involved.
- `bulkRecipientPlan_(sheet, payload)` turns a UI payload into recipient-builder options + a label; `sendBulkEmails` and `previewBulkRecipients` (the form's "Preview recipients" button) both go through it, so the preview count can't disagree with the send.
- Modes everywhere in this repo: `"dry"` (log Pending rows only) | `"test"` (send to editable test recipients) | `"actual"`. Test never consumes idempotency keys.
- Batching: one email per ≤45 recipients (hard cap 50 recipients/message incl. To+CC+BCC; BCC mode reserves 1 for To-self). Default BCC for privacy.
- **Personalize** option: forces one-email-per-person; `{{name}}` in the body → first name (fallback "there"). Subject is NOT personalized (tracking matches on exact subject).
- Quota: `MailApp.getRemainingDailyQuota()` checked before sends.

## 6. Lifecycle emails (`lifecycle_emailer.gs` + Email Composer)

- Config in `lifecycleEmailerConfig_()`: MODE, `SENDER_NAME` ("My name is …"), `SIGNOFF_NAME` ("Warmly, …"), `TEST_RECIPIENT`, `INCLUDE_MAYBE_RSVPS`, `WELCOME_MAX_PRIOR_ATTENDED` (welcome skips regulars), `REMINDER_WINDOW_DAYS`, `EVENT_TIME`, `DEFAULT_ADDRESS`, timezone note (format dates in script TZ or days shift).
- Templates in `EMAIL_TEMPLATES`: `welcome` (upcoming signups, newcomers), `reminder` (today/tomorrow signups), `missed_you` (last event's no-shows), `follow_up` (last event's attendees). Each has `eventScope` + `audience(cell, config)` predicate + `subject`/`build` functions.
- Core: `runLifecycleEmailer_(templateKey, config, only, eventsOverride)` — events from Row 6/7 (`getAllEventColumns_`), addresses from Schedule col E by date, recipients by cell predicate, idempotent per template+email+eventKey.
- **Email Composer** popup picks event + template + mode; shows per-template audience counts; requires a typed "Sent by" name (overrides BOTH `SENDER_NAME` and `SIGNOFF_NAME` for that send) and, in test mode, an editable test recipient. Send hops to the web app which serves an instant progress page.

## 7. Triggers & menu

- `onOpen` → "Custom Actions" menu (maintenance) + "Email" menu (Composer / Bulk Emailer / No-Email Report). Anything that sends mail belongs on the Email menu.
- Nightly time-driven trigger → `nightlyTrigger()` (order matters: schedule import FIRST so event blocks exist before the EventBrite move).
- Spreadsheet `onFormSubmit` → `onSignupFormSubmit` (create once via `setupSignupFormTrigger()`).

## 8. Development & deployment workflow

**Golden rule: git and the live Apps Script project must never drift.** After
EVERY change that is done and verified, sync BOTH fronts in this order:

1. `git add` + `git commit` (imperative subject, body explains the why).
2. `git push` to origin.
3. `clasp push -f` (uploads `scripts/` to the live Apps Script project).
4. If any **web-app-served** file changed (`email_composer*.gs/html`,
   `bulk_emailer_dialog.html`, `facebook_csv_import*`, anything reached via
   `doGet`/`doPost`): `clasp deploy -i <DEPLOYMENT_ID> -d "<short label>"` so
   the stable `/exec` URL serves the new code. `clasp push` alone only updates
   HEAD, not the pinned deployment.

Before relying on clasp, check it exists (`command -v clasp` or `clasp -v`).
**If clasp is missing or not logged in, STOP and tell the user to install and
authenticate it** — do not silently skip the live push:

```bash
npm install -g @google/clasp   # install
clasp login                    # authorize with the Google account that owns the script
```

(In this environment clasp may live at `~/.local/node/bin/clasp` — add that to
PATH if `command -v clasp` misses it.)

- Local repo ↔ Apps Script via **clasp**: `.clasp.json` = `{ "scriptId": "<the bound script's ID>", "rootDir": "scripts" }`. Push with `clasp push -f`.
- Web app: deployed once; keep its **deployment ID** stable so the `/exec` URL never changes. After code changes: `clasp deploy -i <DEPLOYMENT_ID> -d "<label>"` to point the same URL at the new version. The current deployment ID is hardcoded in `COMPOSER_WEBAPP_URL` and in `.claude/hooks/sync-on-stop.sh`.
- This repo has a Claude Code **Stop hook** (`.claude/hooks/sync-on-stop.sh`): when the tree is clean and ahead of origin it auto `git push`es, `clasp push -f`es (if `scripts/` changed), and redeploys the web app. It never auto-commits.
- Rebuilding from scratch: create the bound Apps Script project from the spreadsheet, `clasp clone <scriptId>` (or set `.clasp.json`), push files, deploy the web app (Execute as: user accessing, Access: anyone), paste the new `/exec` URL into `COMPOSER_WEBAPP_URL`, update the hook's `WEBAPP_DEPLOYMENT_ID`, run `setupSignupFormTrigger()`, create the nightly time trigger for `nightlyTrigger`, and authorize scopes on first run.

## 9. Conventions & gotchas for NEW scripts (follow these or things break)

1. **Never hardcode column letters** for counters/landmarks — find them via Row 5 labels / `findColMarker_`. Event columns move when new events are inserted at O.
2. **1-based vs 0-based**: `getRange` is 1-based; `getValues()` arrays and `COLUMN_INDEX` are 0-based. `findEventColumnByTitle_` returns 0-based; most `findEventColumnByDate_`-style helpers return 1-based. Check each helper's contract.
3. **Bulk reads**: one `getRange(...).getValues()` for the whole area, never per-cell reads in loops (3000+ rows).
4. **Row inserts**: expand collapsed groups first (`expandRowGroupsAtRow_` + `showRows`), insert **two rows below** an event-block title, and process multi-row work **bottom-up** so indices don't shift. Delete rows bottom-up (batch consecutive runs).
5. **New person rows** need: C/D names, A concat formula, counter-column formulas (`FORMULAS.COUNT_ATTENDED/COUNT_RSVP` with the attended column letter), RSVP dropdown validation on event cells, dashes in untouched event columns, `UI_CONSTANTS` font/size/alignment.
6. **Never concatenate J/K/L/M** during merges — the original signup wins (merge skip-set).
7. **Event titles may contain commas** — match titles against Row 7 / known titles, never split on comma.
8. **Emails**: split cells on `/[,;]+/`, validate with `/^[^\s@]+@[^\s@]+\.[^\s@]+$/`, compare lowercase; dedupe with Map/Set keyed by email.
9. **Names**: normalize with `normalizeByStrippingWhiteSpaceAtTheEnd`; fuzzy match = exact, spaceless, and first/last swapped. Split full names: first token = first name, rest = last name.
10. **Phones**: compare digits-only (`normalizePhone_`), store dash format, append new numbers with ", ".
11. **Every sender supports dry/test/actual** and logs to a tracking sheet; idempotency keys must survive re-runs; test runs never consume keys.
12. **HTML dialogs**: inject boot data server-side (escape `<` as `<`); escape ALL user strings with an `esc()` helper; popups never call `google.script.run` (see §4).
13. **Apps Script limits**: 50 recipients/message; daily recipient quota (~100 consumer / 1500 Workspace); 100KB per CacheService value; 6-min execution cap — chunk long jobs and write progress.
14. **`Session.getActiveUser().getEmail()`** works here because the web app executes as USER_ACCESSING with the userinfo.email scope.
15. Sheet writes that pair values + formulas: write values first, then re-set formulas for formula cells (see merge's write-back), and refresh count-formula column letters (`FORMULAS.COUNT_*` embed the attended column letter, which can go stale).
