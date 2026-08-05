/**
 * Lifecycle emailer — personalized emails for every stage of the messaging
 * guidance, one template per stage:
 *
 *   welcome    "Excited to meet you!"        → signups of UPCOMING events
 *   reminder   "Look forward to seeing you!" → signups of events happening
 *                                              today/tomorrow (same-day or
 *                                              day-before nudge)
 *   missed_you "Missed you at our event!"    → people who RSVP'd to the LAST
 *                                              event but didn't attend;
 *                                              invites them to the next one
 *   follow_up  "Thank you for joining us!"   → people who ATTENDED the last
 *                                              event; supports a personal
 *                                              note for individual sends
 *
 * Shared mechanics:
 *  - Events live in Contact List columns O.. (Row 6 = date, Row 7 = title);
 *    addresses come from the Schedule sheet (col E, matched by date).
 *  - Recipients are matched by their RSVP/attendance cell in the event column
 *    (Row 12+), emailed by first name (col C).
 *  - Idempotent: every send is logged to the "Lifecycle Email Tracking" sheet
 *    keyed by template + email + event (title|date). Re-running skips anyone
 *    already "Sent" by an Actual Run. Test/dry runs never consume the key.
 *
 * MODE lives in lifecycleEmailerConfig_(): "dry" logs only, "test" sends
 * real individuals' details to TEST_RECIPIENT, "actual" sends for real.
 */

function lifecycleEmailerConfig_() {
  return {
    MODE: "test", // "dry" (log only) | "test" (send to TEST_RECIPIENT) | "actual"

    SENDER_NAME: "Sina", // used in greetings/sign-offs

    // "test" mode: every email goes here, but keeps the real individual's
    // name and event details so you can proofread exactly what they'd get.
    TEST_RECIPIENT: "wintersina@gmail.com",
    TEST_MAX_PER_EVENT: 1, // cap test sends per event (1 sample each)

    // Which RSVP cells count as "signed up" (welcome + reminder audiences)
    INCLUDE_MAYBE_RSVPS: true,
    INCLUDE_TODAY: true, // treat today's event as upcoming

    // reminder: only nudge for events within this many days (1 = today+tomorrow,
    // per the guidance that same-day / day-before messaging is most effective)
    REMINDER_WINDOW_DAYS: 1,

    // Event details used in bodies
    EVENT_TIME: "6:30 PM – 8:00 PM", // matches the calendar sync window
    DEFAULT_ADDRESS: "",             // fallback when Schedule col E has no location

    TRACKING_SHEET_NAME: "Lifecycle Email Tracking",
    // Dates are parsed into script-timezone midnights (parseEventDate_), so
    // formatting MUST use the same zone or day-of-week/date shift by one.
    TIME_ZONE: Session.getScriptTimeZone()
  };
}

/** ————————————————————————————————————————————————————————
 * Entry points (bulk + test per template, one individual sender)
 * ———————————————————————————————————————————————————————— */

function sendWelcomeEmails()          { runLifecycleEmailer_("welcome", lifecycleEmailerConfig_(), null); }
function sendReminderEmails()         { runLifecycleEmailer_("reminder", lifecycleEmailerConfig_(), null); }
function sendMissedYouEmails()        { runLifecycleEmailer_("missed_you", lifecycleEmailerConfig_(), null); }
function sendAttendeeFollowUpEmails() { runLifecycleEmailer_("follow_up", lifecycleEmailerConfig_(), null); }

/** Test variants: force MODE "test" → one sample per event to TEST_RECIPIENT,
 *  built with a REAL individual's name and details. Never marks anyone. */
function sendTestWelcomeEmail()   { runLifecycleTest_("welcome"); }
function sendTestReminderEmail()  { runLifecycleTest_("reminder"); }
function sendTestMissedYouEmail() { runLifecycleTest_("missed_you"); }
function sendTestFollowUpEmail()  { runLifecycleTest_("follow_up"); }

function runLifecycleTest_(templateKey) {
  var config = lifecycleEmailerConfig_();
  config.MODE = "test";
  runLifecycleEmailer_(templateKey, config, null);
}

/**
 * Sends ONE template email to ONE individual. Edit INDIVIDUAL below, then run.
 * MODE still comes from lifecycleEmailerConfig_() ("test" routes their
 * personalized email to TEST_RECIPIENT). Idempotent like the bulk runs;
 * FORCE_RESEND overrides a prior send.
 */
function sendLifecycleEmailToIndividual() {
  var INDIVIDUAL = {
    TEMPLATE: "welcome",          // "welcome" | "reminder" | "missed_you" | "follow_up"
    EMAIL: "someone@example.com", // the person's address in the Contact List
    TOPIC: "One God, Many Paths", // Row 7 title (upcoming for welcome/reminder, past for missed_you/follow_up)
    PERSONAL_NOTE: "",            // follow_up only: e.g. "I loved what you shared about stoicism."
    FORCE_RESEND: false           // true = resend even if already sent
  };
  runLifecycleEmailer_(INDIVIDUAL.TEMPLATE, lifecycleEmailerConfig_(), INDIVIDUAL);
}

/** ————————————————————————————————————————————————————————
 * Template definitions
 * ———————————————————————————————————————————————————————— */

// eventScope: which event columns the template applies to.
//   "upcoming"       — every event dated today or later
//   "reminderWindow" — upcoming events within REMINDER_WINDOW_DAYS
//   "lastPast"       — the most recent event before today
// audience: predicate over the person's cell in the event column.
var EMAIL_TEMPLATES = {
  welcome: {
    label: "Welcome",
    eventScope: "upcoming",
    audience: function(cell, config) { return rsvpCellIsSignup_(cell, config.INCLUDE_MAYBE_RSVPS); },
    subject: function(ev, ctx, config) { return "Excited to meet you!"; },
    build: buildWelcomeEmailBody_
  },
  reminder: {
    label: "Reminder",
    eventScope: "reminderWindow",
    audience: function(cell, config) { return rsvpCellIsSignup_(cell, config.INCLUDE_MAYBE_RSVPS); },
    subject: function(ev, ctx, config) {
      return isSameDay_(ev.date, ctx.today) ? "Look forward to seeing you today!"
                                            : "Look forward to seeing you " + ev.dayOfWeek + "!";
    },
    build: buildReminderEmailBody_
  },
  missed_you: {
    label: "Missed You",
    eventScope: "lastPast",
    audience: function(cell, config) { return cellIsNoShow_(cell); },
    subject: function(ev, ctx, config) { return "Missed you at our event!"; },
    build: buildMissedYouEmailBody_
  },
  follow_up: {
    label: "Follow Up",
    eventScope: "lastPast",
    audience: function(cell, config) { return cellIsAttended_(cell); },
    subject: function(ev, ctx, config) { return "Thank you for joining us!"; },
    build: buildFollowUpEmailBody_
  }
};

/** ————————————————————————————————————————————————————————
 * Core flow
 * ———————————————————————————————————————————————————————— */
function runLifecycleEmailer_(templateKey, config, only, eventsOverride) {
  var tpl = EMAIL_TEMPLATES[templateKey];
  if (!tpl) {
    throw new Error('Unknown template "' + templateKey + '" (use ' + Object.keys(EMAIL_TEMPLATES).join(" | ") + ")");
  }
  if (!["dry", "test", "actual"].includes(config.MODE)) {
    throw new Error('Unknown MODE "' + config.MODE + '" (use "dry" | "test" | "actual")');
  }

  var [contactSheet, _eventbrite, scheduleSheet] = sheetsByName();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tracking = ensureLifecycleTrackingSheet_(ss, config.TRACKING_SHEET_NAME);

  var all = getAllEventColumns_(contactSheet, config);
  var today = startOfToday_();
  var upcoming = all.filter(function(e) { return config.INCLUDE_TODAY ? e.date >= today : e.date > today; });
  var past = all.filter(function(e) { return e.date < today; });

  var locationByDate = getScheduleLocationMap_(scheduleSheet);
  var ctx = {
    today: today,
    nextEvent: upcoming[0] || null,
    nextAddress: upcoming[0] ? (locationByDate[upcoming[0].dateKey] || config.DEFAULT_ADDRESS) : "",
    personalNote: (only && only.PERSONAL_NOTE) ? String(only.PERSONAL_NOTE).trim() : ""
  };

  // Pick the event columns this template applies to.
  var pool = (tpl.eventScope === "lastPast") ? past : upcoming;
  var events;
  if (eventsOverride && eventsOverride.length) {
    events = eventsOverride; // explicit selection (Email Composer dialog)
  } else if (only && only.TOPIC) {
    var normTopic = normalizeByStrippingWhiteSpaceAtTheEnd(only.TOPIC);
    events = pool.filter(function(e) { return e.normTitle === normTopic; });
    if (events.length === 0) {
      throw new Error('Topic "' + only.TOPIC + '" is not ' +
        (tpl.eventScope === "lastPast" ? "a past" : "an upcoming") + ' event. Available: ' +
        (pool.map(function(e) { return e.title; }).join(", ") || "(none)"));
    }
    events = [events[events.length - 1]]; // most recent match
  } else if (tpl.eventScope === "upcoming") {
    events = upcoming;
  } else if (tpl.eventScope === "reminderWindow") {
    var cutoff = new Date(today.getFullYear(), today.getMonth(), today.getDate() + config.REMINDER_WINDOW_DAYS);
    events = upcoming.filter(function(e) { return e.date <= cutoff; });
  } else { // lastPast
    events = past.length ? [past[past.length - 1]] : [];
  }

  if (events.length === 0) {
    Logger.log("[%s] No applicable events (scope: %s). Nothing to send.", tpl.label, tpl.eventScope);
    return { sent: 0, skipped: 0, failed: 0, planned: 0, noEvents: true };
  }
  Logger.log("[%s] Events: %s", tpl.label,
    events.map(function(e) { return e.title + " (" + e.dateStr + ")"; }).join(" | "));

  var alreadySent = buildLifecycleSentSet_(tracking);
  var totals = { sent: 0, skipped: 0, failed: 0, planned: 0 };

  events.forEach(function(ev) {
    var recipients = collectEventAudience_(contactSheet, ev.col0, tpl, config);

    if (only) {
      var target = findRecipientOrContact_(contactSheet, recipients, only.EMAIL);
      if (!target) {
        throw new Error('No Contact List row found with email "' + only.EMAIL + '". Aborting.');
      }
      if (!target.matchedAudience) {
        Logger.log('Note: %s does not match the %s audience for "%s" — sending anyway (explicitly targeted).',
          only.EMAIL, tpl.label, ev.title);
      }
      recipients = [target];
    }

    if (recipients.length === 0) {
      Logger.log('[%s] No eligible recipients for "%s" (%s).', tpl.label, ev.title, ev.dateStr);
      return;
    }

    var address = locationByDate[ev.dateKey] || config.DEFAULT_ADDRESS;
    if (!address && tpl.eventScope !== "lastPast") {
      Logger.log('Warning: no address for "%s" (%s) — Schedule col E is empty and DEFAULT_ADDRESS is blank.', ev.title, ev.dateStr);
    }
    var evCtx = Object.assign({ address: address }, ctx);

    var subject = tpl.subject(ev, evCtx, config);
    var testSentForEvent = 0;

    recipients.forEach(function(person) {
      var key = lifecycleKey_(templateKey, person.email, ev.eventKey);
      var forceResend = !!(only && only.FORCE_RESEND);

      if (config.MODE !== "test" && !forceResend && alreadySent.has(key)) {
        Logger.log("[%s] Skip (already sent): %s for %s", tpl.label, person.email, ev.eventKey);
        totals.skipped++;
        return;
      }

      var body = tpl.build(person.firstName, ev, evCtx, config);

      if (config.MODE === "dry") {
        Logger.log('[%s] Dry run: would email %s <%s> for "%s" (%s).', tpl.label, person.firstName, person.email, ev.title, ev.dateStr);
        appendLifecycleTracking_(tracking, person.email, ev.eventKey, templateKey, "Pending", "Dry Run", person.firstName, "", subject, "");
        totals.planned++;
        return;
      }

      if (config.MODE === "test") {
        if (testSentForEvent >= config.TEST_MAX_PER_EVENT) return;
        var testRes = safeSendEmail_(config.TEST_RECIPIENT, "[TEST] " + subject, body, null);
        testSentForEvent++;
        if (testRes.ok) {
          Logger.log('[%s] Test email sent to %s using %s\'s details for "%s".', tpl.label, config.TEST_RECIPIENT, person.firstName, ev.title);
          appendLifecycleTracking_(tracking, config.TEST_RECIPIENT, ev.eventKey, templateKey, "Sent", "Test Run", person.firstName, "", subject, person.email);
          totals.sent++;
        } else {
          Logger.log("[%s] Test send failed: %s", tpl.label, testRes.error);
          appendLifecycleTracking_(tracking, config.TEST_RECIPIENT, ev.eventKey, templateKey, "Failed", "Test Run", person.firstName, testRes.error, subject, person.email);
          totals.failed++;
        }
        return;
      }

      // actual
      var res = safeSendEmail_(person.email, subject, body, null);
      if (res.ok) {
        Logger.log('[%s] Sent to %s <%s> for "%s" (%s).', tpl.label, person.firstName, person.email, ev.title, ev.dateStr);
        appendLifecycleTracking_(tracking, person.email, ev.eventKey, templateKey, "Sent", "Actual Run", person.firstName, "", subject, "");
        alreadySent.add(key); // guard against duplicate rows in the same run
        totals.sent++;
      } else {
        Logger.log("[%s] Failed to send to %s: %s", tpl.label, person.email, res.error);
        appendLifecycleTracking_(tracking, person.email, ev.eventKey, templateKey, "Failed", "Actual Run", person.firstName, res.error, subject, "");
        totals.failed++;
      }
    });
  });

  Logger.log("[%s] Done. Mode=%s — sent: %s, skipped (already sent): %s, failed: %s, planned (dry): %s",
    tpl.label, config.MODE, totals.sent, totals.skipped, totals.failed, totals.planned);
  return totals;
}

/** ————————————————————————————————————————————————————————
 * Events + audiences
 * ———————————————————————————————————————————————————————— */

function startOfToday_() {
  var now = new Date();
  return new Date(now.getFullYear(), now.getMonth(), now.getDate());
}

function isSameDay_(a, b) {
  return a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth() && a.getDate() === b.getDate();
}

/**
 * Every event column with a title and parseable date, soonest first, as
 * { col0, col1, title, normTitle, date, dateKey, eventKey, dayOfWeek, dateStr }.
 */
function getAllEventColumns_(contactSheet, config) {
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

  var events = [];
  for (var i = 0; i < numCols; i++) {
    var title = titles[i] ? String(titles[i]).trim() : "";
    if (!title) continue;

    var eventDate = parseEventDate_(dates[i]);
    if (!eventDate) continue;

    var key = dateKey_(eventDate);
    events.push({
      col0: startCol - 1 + i, // 0-based index into row arrays
      col1: startCol + i,     // 1-based sheet column
      title: title,
      normTitle: normalizeByStrippingWhiteSpaceAtTheEnd(title),
      date: eventDate,
      dateKey: key,
      eventKey: title + "|" + key, // same topic on a new date is a new send
      dayOfWeek: Utilities.formatDate(eventDate, config.TIME_ZONE, "EEEE"),
      dateStr: Utilities.formatDate(eventDate, config.TIME_ZONE, "MMMM d, yyyy")
    });
  }

  events.sort(function(a, b) { return a.date - b.date; });
  return events;
}

/**
 * Collects unique recipients for one event column from the data rows (Row 12+),
 * keeping rows whose event cell matches the template's audience predicate.
 * Returns [{email, firstName, matchedAudience}].
 */
function collectEventAudience_(contactSheet, eventCol0, tpl, config) {
  var lastRow = contactSheet.getLastRow();
  if (lastRow < ROW_NUMBERS.ROW_12) return [];

  var data = contactSheet.getRange(ROW_NUMBERS.ROW_12, 1, lastRow - ROW_NUMBERS.ROW_12 + 1, contactSheet.getLastColumn()).getValues();
  var emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  var seen = new Set();
  var out = [];

  for (var r = 0; r < data.length; r++) {
    if (!tpl.audience(data[r][eventCol0], config)) continue;

    var rawEmail = (data[r][COLUMN_INDEX.EMAIL] || "").toString().trim();
    if (!rawEmail) continue;
    // Multi-email cells: use the first valid address only (one person, one email)
    var email = rawEmail.split(/[,;]+/)[0].trim();
    if (!emailRegex.test(email)) continue;

    var norm = email.toLowerCase();
    if (seen.has(norm)) continue; // duplicate rows (RSVP + Attended sections)
    seen.add(norm);

    out.push({ email: email, firstName: extractFirstName_(data[r]), matchedAudience: true });
  }
  return out;
}

/** "signed up": rsvp'd yes (or maybe when included). */
function rsvpCellIsSignup_(cellValue, includeMaybe) {
  var v = (cellValue || "").toString().trim().toLowerCase();
  if (!v || v === "-" || v === "--") return false;
  if (v.indexOf("rsvp'd: yes") === 0) return true;
  if (includeMaybe && v.indexOf("rsvp'd: maybe") === 0) return true;
  return false;
}

/** "no-show": rsvp'd yes/maybe but attended: no. Declines (rsvp'd: no) don't count. */
function cellIsNoShow_(cellValue) {
  var v = (cellValue || "").toString().trim().toLowerCase();
  if (!v || v === "-" || v === "--") return false;
  var signedUp = v.indexOf("rsvp'd: yes") === 0 || v.indexOf("rsvp'd: maybe") === 0;
  return signedUp && v.indexOf("attended: no") !== -1;
}

/** "attended": any cell ending in attended: yes, regardless of RSVP. */
function cellIsAttended_(cellValue) {
  var v = (cellValue || "").toString().trim().toLowerCase();
  if (!v || v === "-" || v === "--") return false;
  return v.indexOf("attended: yes") !== -1;
}

/** First name from column C, falling back to the first word of column A. */
function extractFirstName_(row) {
  var first = (row[COLUMN_INDEX.FIRST_NAME] || "").toString().trim();
  if (first) return first;
  var full = (row[COLUMN_INDEX.FULL_NAME_KEY] || "").toString().trim();
  return full ? full.split(/\s+/)[0] : "there";
}

/**
 * For the individual flow: prefer the already-collected recipient (matched the
 * template's audience); otherwise look the person up anywhere in the Contact
 * List so we still know their name. Returns {email, firstName, matchedAudience}
 * or null.
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
    return { email: email.trim(), firstName: extractFirstName_(data[r]), matchedAudience: false };
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
    // Placeholder values ("TBD", "OFF", "TOPIC") are not real addresses.
    if (location && SCHEDULE_SHEET_CONSTANTS.SKIP_WORDS.indexOf(location.toUpperCase()) !== -1) location = "";
    if (!map[key] && location) map[key] = location;
  }
  return map;
}

/** ————————————————————————————————————————————————————————
 * Message bodies (one builder per template)
 * ———————————————————————————————————————————————————————— */

/** Shared: "Day, Date / Time / Title / Address" details block. */
function eventDetailsBlock_(ev, address, config) {
  return ev.dayOfWeek + ", " + ev.dateStr + "\n" + config.EVENT_TIME + "\n" + ev.title + (address ? "\n" + address : "");
}

/** Shared: paragraphs → {text, html}. Empty/null paragraphs are dropped. */
function paragraphsToBody_(paragraphs) {
  var kept = paragraphs.filter(function(p) { return p !== null && p !== undefined && p !== ""; });
  var text = kept.join("\n\n");
  var html = kept.map(function(p) {
    return "<p>" + escapeHtml_(p).replace(/\n/g, "<br>") + "</p>";
  }).join("\n");
  return { text: text, html: html };
}

var ALL_AGES_LINE = "Although events are designed for mature interchange, people of all ages are welcome to contribute perspective.";

/** Template: Initial Response to RSVP. */
function buildWelcomeEmailBody_(firstName, ev, ctx, config) {
  return paragraphsToBody_([
    "Hi " + firstName + "!",

    'So glad you signed up for our upcoming program "' + ev.title + '". My name is ' + config.SENDER_NAME +
      " and I just wanted to confirm that we will be meeting on " + ev.dayOfWeek + ", " + ev.dateStr +
      " at " + config.EVENT_TIME +
      (ctx.address ? ", and have included the address below." : " — I'll follow up with the address once it's confirmed.") +
      " Please let me know if you have any questions.",

    eventDetailsBlock_(ev, ctx.address, config),

    ALL_AGES_LINE,

    "Looking forward to meeting you on " + ev.dayOfWeek + "!",

    "Warmly,\n" + config.SENDER_NAME
  ]);
}

/** Template: Reminder message to RSVPs close to the event date. */
function buildReminderEmailBody_(firstName, ev, ctx, config) {
  var when = isSameDay_(ev.date, ctx.today) ? "today" : "on " + ev.dayOfWeek + ", " + ev.dateStr;
  return paragraphsToBody_([
    "Hi " + firstName + "!",

    "Look forward to seeing you at our upcoming gathering " + when + ". Looks like we have a really nice group " +
      "signed up, and I'm looking forward to meeting everyone. Please do let me know if you have any questions whatsoever!",

    "Here are the event details:",

    eventDetailsBlock_(ev, ctx.address, config),

    ALL_AGES_LINE,

    "See you soon!\n" + config.SENDER_NAME
  ]);
}

/** Template: Post-event message to signups who didn't attend. */
function buildMissedYouEmailBody_(firstName, ev, ctx, config) {
  var next = ctx.nextEvent;
  var invite = next
    ? 'We will be hosting our next gathering, "' + next.title + '", on ' + next.dayOfWeek + ", " + next.dateStr +
      " at " + config.EVENT_TIME + ", and would love to have you join us if you are free. It should be a rich " +
      "conversation, and an opportunity to connect with a diverse group of people."
    : "We would love to have you join us at a future gathering — I'll be sure to share the details for the next one soon.";

  return paragraphsToBody_([
    "Hi " + firstName + ",",

    'Sorry you weren\'t able to attend our program "' + ev.title + '"!',

    invite,

    "Please don't hesitate to let me know if you have any questions.",

    next ? eventDetailsBlock_(next, ctx.nextAddress, config) : "",

    "Warmly,\n" + config.SENDER_NAME
  ]);
}

/** Template: Post-event thank-you for attendees (personal note encouraged). */
function buildFollowUpEmailBody_(firstName, ev, ctx, config) {
  var next = ctx.nextEvent;
  return paragraphsToBody_([
    "Hi " + firstName + "!",

    'Thank you so much for joining us for "' + ev.title + '" — it was great to have you, and I really enjoyed ' +
      "the perspective you brought to the conversation.",

    ctx.personalNote, // dropped when empty

    next
      ? 'Our next gathering, "' + next.title + '", is on ' + next.dayOfWeek + ", " + next.dateStr + " at " +
        config.EVENT_TIME + " — I'd love to see you there if you're free."
      : "",

    next ? eventDetailsBlock_(next, ctx.nextAddress, config) : "",

    "Please don't hesitate to reach out if you have any questions, or if there's anything you'd like to explore further.",

    "Warmly,\n" + config.SENDER_NAME
  ]);
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
var LIFECYCLE_TRACKING_HEADER = ["Email", "Event Key", "Template", "Sent Status", "Run Type", "Name", "Timestamp", "Error", "Subject", "Intended Recipient"];

function ensureLifecycleTrackingSheet_(ss, name) {
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(LIFECYCLE_TRACKING_HEADER);
    return sh;
  }
  var firstCell = sh.getRange(1, 1).getValue();
  if (firstCell !== "Email") {
    sh.insertRows(1, 1);
    sh.getRange(1, 1, 1, LIFECYCLE_TRACKING_HEADER.length).setValues([LIFECYCLE_TRACKING_HEADER]);
  }
  return sh;
}

/** Key that makes a send idempotent: one email per template per person per event. */
function lifecycleKey_(templateKey, email, eventKey) {
  return templateKey + "||" + email.toString().trim().toLowerCase() + "||" + eventKey;
}

/**
 * Set of lifecycleKey_ values already Sent by an ACTUAL run. Test and dry rows
 * never block a future real send.
 */
function buildLifecycleSentSet_(trackingSheet) {
  var vals = trackingSheet.getDataRange().getValues();
  var sent = new Set();
  for (var r = 1; r < vals.length; r++) {
    var email = (vals[r][0] || "").toString().trim();
    var eventKey = (vals[r][1] || "").toString().trim();
    var template = (vals[r][2] || "").toString().trim() || "welcome"; // legacy rows had no Template column
    var status = (vals[r][3] || "").toString().trim();
    var runType = (vals[r][4] || "").toString().trim();
    if (email && eventKey && status === "Sent" && runType === "Actual Run") {
      sent.add(lifecycleKey_(template, email, eventKey));
    }
  }
  return sent;
}

function appendLifecycleTracking_(trackingSheet, email, eventKey, templateKey, status, runType, firstName, err, subject, intendedRecipient) {
  trackingSheet.appendRow([
    email,
    eventKey,
    templateKey,          // "welcome" | "reminder" | "missed_you" | "follow_up"
    status,               // "Sent" | "Failed" | "Pending"
    runType,              // "Dry Run" | "Test Run" | "Actual Run"
    firstName,
    new Date(),
    err || "",
    subject || "",
    intendedRecipient || "" // test runs: the real person the sample was built for
  ]);
}
