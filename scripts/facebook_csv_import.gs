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
 * account with ?authuser=. Performs the import and renders a result page.
 */
function doPost(e) {
  var p = (e && e.parameter) || {};
  if (p.action === "fbimport") return handleFacebookImportPost_(p);
  return HtmlService.createHtmlOutput("Unsupported request.");
}

function handleFacebookImportPost_(p) {
  var ok, message;
  try {
    message = importFacebookGuests(JSON.parse(p.payload));
    ok = true;
  } catch (err) {
    message = (err && err.message) ? err.message : String(err);
    ok = false;
  }
  var runner = Session.getActiveUser().getEmail() || "(unknown account)";
  var html =
    '<div style="font-family:Roboto,Arial,sans-serif;font-size:14px;max-width:560px;margin:48px auto;padding:0 16px">' +
    '<h2 style="color:' + (ok ? "#188038" : "#c5221f") + ';margin-bottom:8px">' +
    (ok ? "✓ Facebook CSV Import — done" : "✗ Facebook CSV Import — failed") + "</h2>" +
    "<p>" + escapeHtml_(message) + "</p>" +
    '<p style="color:#5f6368;font-size:12px">Ran as ' + escapeHtml_(runner) +
    ". You can close this tab.</p>" +
    "</div>";
  return HtmlService.createHtmlOutput(html).setTitle("Facebook CSV Import — result");
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

  var startRow = Math.max(eventbriteSheet.getLastRow() + 1, HELPER_CONSTANTS.FIRST_DATA_ROW);
  eventbriteSheet
    .getRange(startRow, 1, rows.length, HELPER_CONSTANTS.EVENTBRITE_COLUMN_COUNT)
    .setValues(rows);

  var guests = rows.length + (rows.length === 1 ? " guest" : " guests");
  if (!payload.moveNow) {
    return guests + ' staged in "' + SHEET_NAMES.EVENTBRITE + '" (not moved to the Contact List yet).';
  }

  moveRowsFromEventBriteImportToContactList();
  return guests + ' imported into the Contact List under "' + payload.eventTitle + '".';
}
