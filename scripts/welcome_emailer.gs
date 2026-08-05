/**
 * Welcome emailer — personalized "Excited to meet you!" emails to individuals
 * who signed up (RSVP'd) for UPCOMING events.
 *
 * How it decides who to email:
 *  - Upcoming events = Contact List event columns (O .. one before
 *    "# Events Attended") whose Row 6 date is today or later (Row 7 = title).
 *  - Recipients for an event = data rows (Row 12+) whose cell in that event
 *    column reads "rsvp'd: yes ..." (and optionally "rsvp'd: maybe ...").
 *  - The event's address comes from the Schedule sheet (column E, matched by
 *    date), falling back to DEFAULT_ADDRESS.
 *
 * Idempotency:
 *  - Every send is logged to the "Welcome Email Tracking" sheet keyed by
 *    email + event (title|date). A person is emailed AT MOST ONCE per event:
 *    re-running skips anyone with a "Sent" + "Actual Run" row for that event.
 *    Test/dry runs never consume the idempotency key.
 *
 * Entry points (also on the Custom Actions menu):
 *  - sendWelcomeEmails()            — all upcoming events, MODE from config below
 *  - sendTestWelcomeEmail()         — forces MODE "test": sends to TEST_RECIPIENT
 *                                     with a REAL individual's name/details
 *  - sendWelcomeEmailToIndividual() — one person + one topic (config inside it)
 */

function welcomeEmailerConfig_() {
  return {
    MODE: "test", // "dry" (log only) | "test" (send to TEST_RECIPIENT) | "actual"

    SUBJECT: "Excited to meet you!",
    SENDER_NAME: "Sina", // used in the greeting/sign-off

    // "test" mode: every email goes here, but keeps the real individual's
    // name and event details so you can proofread exactly what they'd get.
    TEST_RECIPIENT: "wintersina@gmail.com",
    TEST_MAX_PER_EVENT: 1, // cap test sends per upcoming event (1 sample each)

    // Which RSVP cells count as "signed up"
    INCLUDE_MAYBE_RSVPS: true, // also welcome "rsvp'd: maybe" rows
    INCLUDE_TODAY: true,       // treat today's event as upcoming

    // Event details used in the body
    EVENT_TIME: "6:30 PM – 8:00 PM", // matches the calendar sync window
    DEFAULT_ADDRESS: "",             // fallback when Schedule col E has no location

    TRACKING_SHEET_NAME: "Welcome Email Tracking",
    TIME_ZONE: "America/Chicago"
  };
}

/** Sends welcome emails for every upcoming event, per welcomeEmailerConfig_(). */
function sendWelcomeEmails() {
  runWelcomeEmailer_(welcomeEmailerConfig_(), null);
}

/**
 * Test entry point: forces MODE "test" so every upcoming event sends ONE sample
 * email to TEST_RECIPIENT (wintersina@gmail.com) carrying a real individual's
 * first name and that event's real details. Never marks anyone as welcomed.
 */
function sendTestWelcomeEmail() {
  var config = welcomeEmailerConfig_();
  config.MODE = "test";
  runWelcomeEmailer_(config, null);
}

/**
 * Sends the welcome email to ONE individual for ONE upcoming topic.
 * Edit INDIVIDUAL below, then run. MODE still comes from welcomeEmailerConfig_()
 * ("test" sends their personalized email to TEST_RECIPIENT instead).
 * Idempotent like the bulk run; set FORCE_RESEND to override a prior send.
 */
function sendWelcomeEmailToIndividual() {
  var INDIVIDUAL = {
    EMAIL: "someone@example.com",          // the person's address in the Contact List
    TOPIC: "One God, Many Paths",          // Row 7 title of an UPCOMING event
    FORCE_RESEND: false                    // true = resend even if already welcomed
  };
  runWelcomeEmailer_(welcomeEmailerConfig_(), INDIVIDUAL);
}

/** ————————————————————————————————————————————————————————
 * Core flow
 * ———————————————————————————————————————————————————————— */
function runWelcomeEmailer_(config, only) {
  if (!["dry", "test", "actual"].includes(config.MODE)) {
    throw new Error('Unknown MODE "' + config.MODE + '" (use "dry" | "test" | "actual")');
  }

  var [contactSheet, _eventbrite, scheduleSheet] = sheetsByName();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tracking = ensureWelcomeTrackingSheet_(ss, config.TRACKING_SHEET_NAME);

  var events = getUpcomingEventColumns_(contactSheet, config);
  if (events.length === 0) {
    Logger.log("No upcoming events found (Row 6 dates on/after today). Nothing to send.");
    return;
  }
  Logger.log("Upcoming events: " + events.map(function(e) { return e.title + " (" + e.dateStr + ")"; }).join(" | "));

  // If targeting one individual, narrow to their topic's event column.
  if (only) {
    var normTopic = normalizeByStrippingWhiteSpaceAtTheEnd(only.TOPIC);
    events = events.filter(function(e) { return e.normTitle === normTopic; });
    if (events.length === 0) {
      throw new Error('Topic "' + only.TOPIC + '" is not an upcoming event. Upcoming topics: ' +
        getUpcomingEventColumns_(contactSheet, config).map(function(e) { return e.title; }).join(", "));
    }
  }

  var locationByDate = getScheduleLocationMap_(scheduleSheet);
  var alreadyWelcomed = buildWelcomeSentSet_(tracking);

  var totals = { sent: 0, skipped: 0, failed: 0, planned: 0 };

  events.forEach(function(ev) {
    var recipients = collectEventRsvps_(contactSheet, ev.col0, config.INCLUDE_MAYBE_RSVPS);

    if (only) {
      var target = findRecipientOrContact_(contactSheet, recipients, only.EMAIL);
      if (!target) {
        throw new Error('No Contact List row found with email "' + only.EMAIL + '". Aborting.');
      }
      if (!target.hasRsvp) {
        Logger.log('Note: %s has no RSVP recorded for "%s" — sending anyway (explicitly targeted).', only.EMAIL, ev.title);
      }
      recipients = [target];
    }

    if (recipients.length === 0) {
      Logger.log('No eligible signups for "%s" (%s).', ev.title, ev.dateStr);
      return;
    }

    var address = locationByDate[ev.dateKey] || config.DEFAULT_ADDRESS;
    if (!address) {
      Logger.log('Warning: no address found for "%s" (%s) — Schedule col E is empty and DEFAULT_ADDRESS is blank.', ev.title, ev.dateStr);
    }

    var testSentForEvent = 0;

    recipients.forEach(function(person) {
      var key = welcomeKey_(person.email, ev.eventKey);
      var forceResend = !!(only && only.FORCE_RESEND);

      if (config.MODE !== "test" && !forceResend && alreadyWelcomed.has(key)) {
        Logger.log("Skip (already welcomed): %s for %s", person.email, ev.eventKey);
        totals.skipped++;
        return;
      }

      var body = buildWelcomeEmailBody_(person.firstName, ev, address, config);

      if (config.MODE === "dry") {
        Logger.log('Dry run: would welcome %s <%s> for "%s" (%s).', person.firstName, person.email, ev.title, ev.dateStr);
        appendWelcomeTracking_(tracking, person.email, ev.eventKey, "Pending", "Dry Run", person.firstName, "", config.SUBJECT, "");
        totals.planned++;
        return;
      }

      if (config.MODE === "test") {
        if (testSentForEvent >= config.TEST_MAX_PER_EVENT) return;
        var testRes = safeSendEmail_(config.TEST_RECIPIENT, "[TEST] " + config.SUBJECT, body, null);
        testSentForEvent++;
        if (testRes.ok) {
          Logger.log('Test email sent to %s using %s\'s details for "%s".', config.TEST_RECIPIENT, person.firstName, ev.title);
          appendWelcomeTracking_(tracking, config.TEST_RECIPIENT, ev.eventKey, "Sent", "Test Run", person.firstName, "", config.SUBJECT, person.email);
          totals.sent++;
        } else {
          Logger.log("Test send failed: %s", testRes.error);
          appendWelcomeTracking_(tracking, config.TEST_RECIPIENT, ev.eventKey, "Failed", "Test Run", person.firstName, testRes.error, config.SUBJECT, person.email);
          totals.failed++;
        }
        return;
      }

      // actual
      var res = safeSendEmail_(person.email, config.SUBJECT, body, null);
      if (res.ok) {
        Logger.log('Welcomed %s <%s> for "%s" (%s).', person.firstName, person.email, ev.title, ev.dateStr);
        appendWelcomeTracking_(tracking, person.email, ev.eventKey, "Sent", "Actual Run", person.firstName, "", config.SUBJECT, "");
        alreadyWelcomed.add(key); // guard against duplicate rows in the same run
        totals.sent++;
      } else {
        Logger.log("Failed to welcome %s: %s", person.email, res.error);
        appendWelcomeTracking_(tracking, person.email, ev.eventKey, "Failed", "Actual Run", person.firstName, res.error, config.SUBJECT, "");
        totals.failed++;
      }
    });
  });

  Logger.log("Welcome emailer done. Mode=%s — sent: %s, skipped (already welcomed): %s, failed: %s, planned (dry): %s",
    config.MODE, totals.sent, totals.skipped, totals.failed, totals.planned);
}

/** ————————————————————————————————————————————————————————
 * Upcoming events + recipients
 * ———————————————————————————————————————————————————————— */

/**
 * Returns upcoming event columns as
 * { col0, col1, title, normTitle, date, dateKey, eventKey, dayOfWeek, dateStr }.
 * Upcoming = Row 6 date strictly after today, or today too when INCLUDE_TODAY.
 */
function getUpcomingEventColumns_(contactSheet, config) {
  var attendedCol1 = findColMarker_(contactSheet, MARKER_KEYS.EVENTS_ATTENDED, COL_CONSTANTS.EVENTS_ATTENDED);
  if (attendedCol1 === -1) {
    throw new Error('Could not find the "' + COL_CONSTANTS.EVENTS_ATTENDED + '" column (Row 5).');
  }

  var startCol = HELPER_CONSTANTS.EVENT_NAMES_START_COL; // O = 15
  var endCol = attendedCol1 - 1;
  if (endCol < startCol) return [];

  var numCols = endCol - startCol + 1;
  var dates = contactSheet.getRange(ROW_NUMBERS.ROW_6, startCol, 1, numCols).getValues()[0];
  var titles = contactSheet.getRange(ROW_NUMBERS.ROW_7, startCol, 1, numCols).getValues()[0];

  var now = new Date();
  var today = new Date(now.getFullYear(), now.getMonth(), now.getDate());

  var events = [];
  for (var i = 0; i < numCols; i++) {
    var title = titles[i] ? String(titles[i]).trim() : "";
    if (!title) continue;

    var eventDate = parseEventDate_(dates[i]);
    if (!eventDate) continue;

    var isUpcoming = config.INCLUDE_TODAY ? (eventDate >= today) : (eventDate > today);
    if (!isUpcoming) continue;

    var key = dateKey_(eventDate);
    events.push({
      col0: startCol - 1 + i, // 0-based index into row arrays
      col1: startCol + i,     // 1-based sheet column
      title: title,
      normTitle: normalizeByStrippingWhiteSpaceAtTheEnd(title),
      date: eventDate,
      dateKey: key,
      eventKey: title + "|" + key, // same topic on a new date welcomes again
      dayOfWeek: Utilities.formatDate(eventDate, config.TIME_ZONE, "EEEE"),
      dateStr: Utilities.formatDate(eventDate, config.TIME_ZONE, "MMMM d, yyyy")
    });
  }

  // Soonest first
  events.sort(function(a, b) { return a.date - b.date; });
  return events;
}

/**
 * Collects unique signups for one event column from the data rows (Row 12+).
 * A row qualifies when its event cell starts with "rsvp'd: yes" (or
 * "rsvp'd: maybe" when includeMaybe). Returns [{email, firstName, hasRsvp}].
 */
function collectEventRsvps_(contactSheet, eventCol0, includeMaybe) {
  var lastRow = contactSheet.getLastRow();
  if (lastRow < ROW_NUMBERS.ROW_12) return [];

  var data = contactSheet.getRange(ROW_NUMBERS.ROW_12, 1, lastRow - ROW_NUMBERS.ROW_12 + 1, contactSheet.getLastColumn()).getValues();
  var emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  var seen = new Set();
  var out = [];

  for (var r = 0; r < data.length; r++) {
    if (!rsvpCellIsSignup_(data[r][eventCol0], includeMaybe)) continue;

    var rawEmail = (data[r][COLUMN_INDEX.EMAIL] || "").toString().trim();
    if (!rawEmail) continue;
    // Multi-email cells: welcome the first valid address only (one person, one email)
    var email = rawEmail.split(/[,;]+/)[0].trim();
    if (!emailRegex.test(email)) continue;

    var norm = email.toLowerCase();
    if (seen.has(norm)) continue; // duplicate rows (RSVP + Attended sections)
    seen.add(norm);

    out.push({ email: email, firstName: extractFirstName_(data[r]), hasRsvp: true });
  }
  return out;
}

/** True when an event cell means "this person signed up". */
function rsvpCellIsSignup_(cellValue, includeMaybe) {
  var v = (cellValue || "").toString().trim().toLowerCase();
  if (!v || v === "-" || v === "--") return false;
  if (v.indexOf("rsvp'd: yes") === 0) return true;
  if (includeMaybe && v.indexOf("rsvp'd: maybe") === 0) return true;
  return false;
}

/** First name from column C, falling back to the first word of column A. */
function extractFirstName_(row) {
  var first = (row[COLUMN_INDEX.FIRST_NAME] || "").toString().trim();
  if (first) return first;
  var full = (row[COLUMN_INDEX.FULL_NAME_KEY] || "").toString().trim();
  return full ? full.split(/\s+/)[0] : "there";
}

/**
 * For the individual flow: prefer the already-collected recipient (has an RSVP
 * for the event); otherwise look the person up anywhere in the Contact List so
 * we still know their name. Returns {email, firstName, hasRsvp} or null.
 */
function findRecipientOrContact_(contactSheet, recipients, email) {
  var norm = (email || "").toString().trim().toLowerCase();
  for (var i = 0; i < recipients.length; i++) {
    if (recipients[i].email.toLowerCase() === norm) return recipients[i];
  }

  var lastRow = contactSheet.getLastRow();
  if (lastRow < ROW_NUMBERS.ROW_12) return null;
  var data = contactSheet.getRange(ROW_NUMBERS.ROW_12, 1, lastRow - ROW_NUMBERS.ROW_12 + 1, contactSheet.getLastColumn()).getValues();

  for (var r = 0; r < data.length; r++) {
    var cell = (data[r][COLUMN_INDEX.EMAIL] || "").toString().toLowerCase();
    if (cell.indexOf(norm) === -1) continue;
    var hit = cell.split(/[,;]+/).some(function(e) { return e.trim() === norm; });
    if (!hit) continue;
    return { email: email.trim(), firstName: extractFirstName_(data[r]), hasRsvp: false };
  }
  return null;
}

/** Schedule sheet → { "yyyy-MM-dd": location } from columns C (date) and E (location). */
function getScheduleLocationMap_(scheduleSheet) {
  var map = {};
  var lastRow = scheduleSheet.getLastRow();
  if (lastRow < 2) return map;

  var data = scheduleSheet.getRange(2, 3, lastRow - 1, 3).getValues(); // C: date, D: topic, E: location
  for (var i = 0; i < data.length; i++) {
    var d = parseEventDate_(data[i][0]);
    if (!d) continue;
    var key = dateKey_(d);
    var location = data[i][2] ? String(data[i][2]).trim() : "";
    if (!map[key] && location) map[key] = location;
  }
  return map;
}

/** ————————————————————————————————————————————————————————
 * Message body (template: "Initial Response to RSVP")
 * ———————————————————————————————————————————————————————— */
function buildWelcomeEmailBody_(firstName, ev, address, config) {
  var paragraphs = [
    "Hi " + firstName + "!",

    'So glad you signed up for our upcoming program "' + ev.title + '". My name is ' + config.SENDER_NAME +
      " and I just wanted to confirm that we will be meeting on " + ev.dayOfWeek + ", " + ev.dateStr +
      " at " + config.EVENT_TIME + ", and have included the address below. Please let me know if you have any questions.",

    ev.dayOfWeek + ", " + ev.dateStr + "\n" + config.EVENT_TIME + "\n" + ev.title + (address ? "\n" + address : ""),

    "Although events are designed for mature interchange, people of all ages are welcome to contribute perspective.",

    "Looking forward to meeting you on " + ev.dayOfWeek + "!",

    "Warmly,\n" + config.SENDER_NAME
  ];

  var text = paragraphs.join("\n\n");
  var html = paragraphs.map(function(p) {
    return "<p>" + escapeHtml_(p).replace(/\n/g, "<br>") + "</p>";
  }).join("\n");

  return { text: text, html: html };
}

function escapeHtml_(s) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

/** ————————————————————————————————————————————————————————
 * Tracking / idempotency
 * ———————————————————————————————————————————————————————— */
var WELCOME_TRACKING_HEADER = ["Email", "Event Key", "Sent Status", "Run Type", "Name", "Timestamp", "Error", "Subject", "Intended Recipient"];

function ensureWelcomeTrackingSheet_(ss, name) {
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(WELCOME_TRACKING_HEADER);
    return sh;
  }
  var firstCell = sh.getRange(1, 1).getValue();
  if (firstCell !== "Email") {
    sh.insertRows(1, 1);
    sh.getRange(1, 1, 1, WELCOME_TRACKING_HEADER.length).setValues([WELCOME_TRACKING_HEADER]);
  }
  return sh;
}

/** Key that makes a send idempotent: one welcome per person per event. */
function welcomeKey_(email, eventKey) {
  return email.toString().trim().toLowerCase() + "||" + eventKey;
}

/**
 * Set of welcomeKey_ values already Sent by an ACTUAL run. Test and dry rows
 * never block a future real send.
 */
function buildWelcomeSentSet_(trackingSheet) {
  var vals = trackingSheet.getDataRange().getValues();
  var sent = new Set();
  for (var r = 1; r < vals.length; r++) {
    var email = (vals[r][0] || "").toString().trim();
    var eventKey = (vals[r][1] || "").toString().trim();
    var status = (vals[r][2] || "").toString().trim();
    var runType = (vals[r][3] || "").toString().trim();
    if (email && eventKey && status === "Sent" && runType === "Actual Run") {
      sent.add(welcomeKey_(email, eventKey));
    }
  }
  return sent;
}

function appendWelcomeTracking_(trackingSheet, email, eventKey, status, runType, firstName, err, subject, intendedRecipient) {
  trackingSheet.appendRow([
    email,
    eventKey,
    status,               // "Sent" | "Failed" | "Pending"
    runType,              // "Dry Run" | "Test Run" | "Actual Run"
    firstName,
    new Date(),
    err || "",
    subject || "",
    intendedRecipient || "" // test runs: the real person the sample was built for
  ]);
}
