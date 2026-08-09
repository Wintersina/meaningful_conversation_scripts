/**
 * Facebook CSV Import dialog — a popup (Custom Actions → "Facebook CSV Import…")
 * that takes a dragged-in guest-list CSV exported from a Facebook event, keeps
 * only the guests marked "Going", and stages them as EventBrite-shaped rows so
 * the regular EventBrite import carries them into the Contact List. No separate
 * FaceBook Import pipeline anymore.
 *
 * Server side of scripts/facebook_csv_import_dialog.html. Rows land in the
 * "EventBrite Import" tab as A..L = [full name, First, Last, -, -, -, -, -, -,
 * "FB", event title, event code], which is the exact shape
 * moveRowsFromEventBriteImportToContactList expects (it copies B..L into the
 * Contact List and marks everyone "rsvp'd: yes" — hence Going-only). By default
 * that move runs immediately, so the guests land in the Contact List in one step.
 */

/**
 * uiMode: "dialog" (in-sheet popup — Import hops through the web app, because
 * iframe google.script.run calls bind to the browser's default Google session)
 * or "webapp" (full tab — google.script.run is session-safe there, call direct).
 */
function buildFacebookCsvImportHtml_(uiMode) {
  // Same pattern as the Email Composer: inject the boot data at render time
  // instead of fetching it via google.script.run (multi-account sessions).
  var t = HtmlService.createTemplateFromFile("facebook_csv_import_dialog");
  var data = getFacebookCsvImportData();
  data.webappUrl = COMPOSER_WEBAPP_URL;
  data.sendAccount = COMPOSER_SEND_ACCOUNT;
  t.bootData = JSON.stringify(data).replace(/</g, "\\u003c");
  t.uiMode = uiMode;
  return t.evaluate();
}

function showFacebookCsvImportDialog() {
  var html = buildFacebookCsvImportHtml_("dialog").setWidth(500).setHeight(640);
  SpreadsheetApp.getUi().showModalDialog(html, "Facebook CSV Import");
}

/**
 * Web-app POST entry point. The popup's Import hops here via a hidden form
 * POST (guest lists are too large for GET query params), pinned to the team
 * account with ?authuser=. Returns an instant "Importing…" progress page —
 * the page runs the import via google.script.run (session-safe in a full
 * tab) and polls CacheService snapshots, same as the Email Composer send.
 */
function doPost(e) {
  var p = (e && e.parameter) || {};
  if (p.action === "fbimport") return buildFacebookImportProgressPage_(p.payload);
  return HtmlService.createHtmlOutput("Unsupported request.");
}

function buildFacebookImportProgressPage_(payloadStr) {
  var payload;
  try {
    payload = JSON.parse(payloadStr);
  } catch (err) {
    return HtmlService.createHtmlOutput("Bad import payload.").setTitle("Facebook CSV Import");
  }
  payload.progressToken = Utilities.getUuid();

  var t = HtmlService.createTemplateFromFile("facebook_import_progress_page");
  t.payloadJson = JSON.stringify({
    payload: payload,
    account: Session.getActiveUser().getEmail() || "(unknown account)"
  }).replace(/</g, "\\u003c");
  return t.evaluate()
    .setTitle("Facebook CSV Import — importing")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

/**
 * Phase snapshot for the import progress page. Same cache key family the
 * composer/bulk sends use, read by getComposerSendProgress. Best-effort.
 */
function fbImportProgress_(token, processed, total, label, done) {
  if (!token) return;
  try {
    CacheService.getScriptCache().put(
      "composerProgress:" + token,
      JSON.stringify({ total: total, processed: processed, label: label, done: !!done }),
      600
    );
  } catch (e) {}
}

/** Paren-stripping used by the EventBrite move's title matching, mirrored. */
function cleanEventTitle_(title) {
  return normalizeString(String(title || ""))
    .replace(/\s*\(.*?\)\s*/g, "")
    .trim();
}

/**
 * Events offered in the dialog's dropdown (soonest upcoming first, then past,
 * most recent first). Only titles the EventBrite move can actually resolve are
 * offered: it matches the paren-stripped title against Contact List column B
 * (bottom-most block wins), so a title with no column-B block would import
 * into nothing. Duplicate titles collapse to their latest date for the same
 * reason — the move always targets the bottom-most matching block.
 */
function getFacebookCsvImportData() {
  var config = lifecycleEmailerConfig_();
  var [contactSheet] = sheetsByName();
  var all = getAllEventColumns_(contactSheet, config); // sorted by date, ascending
  var today = startOfToday_();

  var contactColB = contactSheet
    .getRange(1, 2, contactSheet.getLastRow(), 1)
    .getValues()
    .flat();

  var byKey = {};
  all.forEach(function(e) {
    var cleaned = cleanEventTitle_(e.title);
    if (!cleaned || contactColB.lastIndexOf(cleaned) === -1) return;
    byKey[cleaned] = { title: e.title, dateStr: e.dateStr, dayOfWeek: e.dayOfWeek, upcoming: e.date >= today };
  });

  var upcoming = [];
  var past = [];
  Object.keys(byKey).forEach(function(key) {
    (byKey[key].upcoming ? upcoming : past).push(byKey[key]);
  });
  past.reverse(); // most recent past event first

  return { upcoming: upcoming, past: past };
}

/**
 * Event code for a title, read from the same 'Event IDs'!B:C mapping the
 * sheet's VLOOKUP header formula uses. Empty string when absent.
 */
function lookupEventCode_(title) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Event IDs");
    if (!sheet) return "";
    var target = String(title || "").trim();
    var rows = sheet.getRange(1, 2, sheet.getLastRow(), 2).getValues(); // B:C
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === target) return String(rows[i][1]).trim();
    }
  } catch (e) {
    Logger.log("Event code lookup failed: %s", e.message);
  }
  return "";
}

/**
 * Imports from the dialog. payload = { guests: [[first, last], ...], eventTitle, moveNow }.
 * Stages EventBrite-shaped rows; when moveNow is set, runs the regular
 * EventBrite move — note that the move sweeps the WHOLE tab, so any real
 * EventBrite rows already sitting there go into the Contact List too.
 * Returns a human-readable summary string shown in the dialog.
 */
function importFacebookGuests(payload) {
  if (!payload || !payload.eventTitle || !Array.isArray(payload.guests) || payload.guests.length === 0) {
    throw new Error("Missing guests or event.");
  }

  var [, eventbriteSheet] = sheetsByName();
  var eventCode = lookupEventCode_(payload.eventTitle);

  var rows = payload.guests
    .map(function(g) {
      var first = String(g[0] || "").trim();
      var last = String(g[1] || "").trim();
      return [
        (first + " " + last).trim(), // A — display only; the move copies B..L
        first,                       // B — First Name
        last,                        // C — Last Name
        "", "", "", "", "", "",      // D..H unknown from Facebook; I — signup date/time
        "FB",                        // J — Original Signup Platform
        payload.eventTitle,          // K — Original Signup Event
        eventCode                    // L — Event Code
      ];
    })
    .filter(function(r) { return r[1]; });
  if (rows.length === 0) throw new Error("No guest names found in the file.");

  var token = payload.progressToken;
  var steps = payload.moveNow ? 2 : 1;
  fbImportProgress_(token, 0, steps, "Staging " + rows.length + " guest(s) in the EventBrite Import tab…");

  var startRow = Math.max(eventbriteSheet.getLastRow() + 1, HELPER_CONSTANTS.FIRST_DATA_ROW);
  eventbriteSheet
    .getRange(startRow, 1, rows.length, HELPER_CONSTANTS.EVENTBRITE_COLUMN_COUNT)
    .setValues(rows);

  var guests = rows.length + (rows.length === 1 ? " guest" : " guests");
  if (!payload.moveNow) {
    fbImportProgress_(token, 1, steps, "Staged.", true);
    return guests + ' staged in "' + SHEET_NAMES.EVENTBRITE + '" (not moved to the Contact List yet).';
  }

  fbImportProgress_(token, 1, steps, "Moving guests into the Contact List (rows, formulas, formatting) — this is the slow part…");
  moveRowsFromEventBriteImportToContactList();
  fbImportProgress_(token, 2, steps, "Done.", true);
  return guests + ' imported into the Contact List under "' + payload.eventTitle + '".';
}
