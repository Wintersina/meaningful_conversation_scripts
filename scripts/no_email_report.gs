/**
 * No-Email Report — everyone on the Contact List (or in one emailing section)
 * who can't be reached by email, so they can be texted or messaged on the
 * platform they signed up through instead.
 *
 * A person lands on the report when NONE of their rows holds a valid address:
 * blank column F, or something in it that isn't a real address ("n/a",
 * "no email", a typo). Rows are grouped by person first — section copies mean
 * the same human appears two or three times, and an address found on any of
 * those rows counts — so nobody is chased who is already reachable.
 *
 * Output: the "No Email Report" sheet, rebuilt on every run, with the phone
 * number, the signup platform (col K) and event (col L) to reach out through,
 * their attendance counts, and which Contact List rows they came from. When
 * everyone has an address the sheet says so explicitly and the count is zero.
 *
 * Entry points: Custom Actions → "No-Email Report…" (whole list) and the
 * "No-email report" button in the Bulk Emailer, which reports on the audience
 * currently selected there.
 */

var NO_EMAIL_REPORT = {
  SHEET_NAME: "No Email Report",
  HEADERS: ["Name", "Phone", "Reach out via", "Signup platform", "Signup event",
            "Signup date", "# Attended", "# RSVP'd", "Issue", "Contact List row(s)"]
};

/** Column K codes → how to actually reach that person. */
var SIGNUP_PLATFORM_OUTREACH = {
  "fb": "Facebook message",
  "facebook": "Facebook message",
  "eb": "Eventbrite message",
  "eventbrite": "Eventbrite message",
  "mu": "Meetup message",
  "meetup": "Meetup message",
  "walk-in": "In person at the next event",
  "walkin": "In person at the next event"
};

/** Trimmed string form of a cell, safe for dates/numbers/blanks. */
function reportCellText_(v) {
  return (v == null) ? "" : String(v).trim();
}

/** Adds a non-empty value to a list, keeping order and dropping duplicates. */
function pushUnique_(list, value) {
  var v = reportCellText_(value);
  if (v && list.indexOf(v) === -1) list.push(v);
}

/**
 * Groups the rows of an audience by person and returns those with no valid
 * email anywhere. Each entry:
 *   { name, phones[], platforms[], events[], signupDate, attended, rsvpd,
 *     rows[], rawEmails[] }
 * Returns { people: [...], peopleScanned: <distinct people in the audience> }.
 */
function collectNoEmailPeople_(contactSheet, audience) {
  var data = contactSheet.getDataRange().getValues();
  var bounds = bulkAudienceBounds_(data, audience || BULK_AUDIENCES.WHOLE);
  var endIdx = Math.min(bounds.end, data.length - 1);
  var emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  var tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();

  var attendedCol1 = findColMarker_(contactSheet, MARKER_KEYS.EVENTS_ATTENDED, COL_CONSTANTS.EVENTS_ATTENDED);
  var rsvpCol1 = findColMarker_(contactSheet, MARKER_KEYS.EVENTS_RSVPD, COL_CONSTANTS.EVENTS_RSVPD);
  var attendedCol0 = (attendedCol1 > 0) ? attendedCol1 - 1 : -1;
  var rsvpCol0 = (rsvpCol1 > 0) ? rsvpCol1 - 1 : -1;

  var people = new Map();

  for (var r = bounds.start; r <= endIdx; r++) {
    var row = data[r];
    var first = reportCellText_(row[COLUMN_INDEX.FIRST_NAME]);
    var last = reportCellText_(row[COLUMN_INDEX.LAST_NAME]);
    // Block titles, "Total RSVP'd" rows and section markers have no first/last
    // name — that's what separates a person row from the sheet's scaffolding.
    if (!first && !last) continue;

    var name = (first + " " + last).trim() || reportCellText_(row[COLUMN_INDEX.FULL_NAME_KEY]);
    var phone = reportCellText_(row[COLUMN_INDEX.PHONE]);
    var key = normalizeByStrippingWhiteSpaceAtTheEnd(name) || normalizePhone_(phone) || ("row:" + r);

    var p = people.get(key);
    if (!p) {
      p = { name: name, phones: [], platforms: [], events: [], signupDate: "",
            attended: "", rsvpd: "", rows: [], rawEmails: [], validEmails: 0 };
      people.set(key, p);
    }

    p.rows.push(r + 1); // 1-based sheet row
    phone.split(/[,;]+/).forEach(function(ph) { pushUnique_(p.phones, formatPhoneWithDashes_(ph.trim())); });
    pushUnique_(p.platforms, row[COLUMN_INDEX.SIGNUP_PLATFORM]);
    pushUnique_(p.events, row[COLUMN_INDEX.SIGNUP_EVENT_TITLE]);

    if (!p.signupDate) {
      var signup = row[COLUMN_INDEX.SIGNUP_DATE_TIME];
      p.signupDate = (signup instanceof Date)
        ? Utilities.formatDate(signup, tz, "MMM d, yyyy")
        : reportCellText_(signup);
    }
    if (attendedCol0 !== -1) p.attended = maxCount_(p.attended, row[attendedCol0]);
    if (rsvpCol0 !== -1) p.rsvpd = maxCount_(p.rsvpd, row[rsvpCol0]);

    reportCellText_(row[COLUMN_INDEX.EMAIL]).split(/[,;]+/).forEach(function(e) {
      var em = e.trim();
      if (!em) return;
      if (emailRegex.test(em)) p.validEmails++;
      else pushUnique_(p.rawEmails, em);
    });
  }

  var unreachable = [];
  people.forEach(function(p) { if (p.validEmails === 0) unreachable.push(p); });

  return { people: unreachable, peopleScanned: people.size };
}

/** Keeps the larger of two counter-cell values (section copies can lag). */
function maxCount_(current, candidate) {
  var n = Number(candidate);
  if (!isFinite(n) || reportCellText_(candidate) === "") return current;
  var c = Number(current);
  return (isFinite(c) && c >= n) ? current : n;
}

/** The best available way to reach someone with no email address. */
function outreachHint_(person) {
  if (person.phones.length) return "Text or call " + person.phones.join(", ");
  for (var i = 0; i < person.platforms.length; i++) {
    var label = SIGNUP_PLATFORM_OUTREACH[person.platforms[i].toLowerCase()];
    if (label) return label;
  }
  if (person.platforms.length) return "Message them on " + person.platforms.join(", ");
  return "No phone or platform on file — ask at the next event";
}

/**
 * (Re)builds the report sheet. Returns the sheet so callers can link to it.
 * With nobody to report, the sheet still gets written — with an explicit
 * "zero" line — so an empty result is never mistaken for a stale run.
 */
function writeNoEmailReportSheet_(ss, people, audience, peopleScanned) {
  var sheet = ss.getSheetByName(NO_EMAIL_REPORT.SHEET_NAME) || ss.insertSheet(NO_EMAIL_REPORT.SHEET_NAME);
  sheet.clear();
  // clear() leaves merged regions in place — break them so the previous run's
  // banner/zero-line merges can't distort this one.
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).breakApart();

  var tz = ss.getSpreadsheetTimeZone();
  var stamp = Utilities.formatDate(new Date(), tz, "MMM d, yyyy 'at' h:mm a");
  var audienceLabel = BULK_AUDIENCE_LABELS[audience] || audience;
  var cols = NO_EMAIL_REPORT.HEADERS.length;

  sheet.getRange(1, 1).setValue(
    "No Email Report — " + audienceLabel + " — " + people.length + " of " + peopleScanned +
    " people have no usable email address (generated " + stamp + ")"
  );
  sheet.getRange(1, 1, 1, cols).merge().setFontWeight("bold").setBackground(UI_CONSTANTS.GRAY_BACKGROUND);

  sheet.getRange(2, 1, 1, cols).setValues([NO_EMAIL_REPORT.HEADERS]).setFontWeight("bold");

  if (people.length === 0) {
    sheet.getRange(3, 1).setValue(
      "Zero — everyone in the " + audienceLabel + " (" + peopleScanned + " people) has a valid email address."
    );
    sheet.getRange(3, 1, 1, cols).merge();
  } else {
    var rows = people.map(function(p) {
      return [
        p.name,
        p.phones.join(", "),
        outreachHint_(p),
        p.platforms.join(", "),
        p.events.join(" / "),
        p.signupDate,
        p.attended,
        p.rsvpd,
        p.rawEmails.length ? ('Unusable address: "' + p.rawEmails.join('", "') + '"') : "No email on file",
        p.rows.join(", ")
      ];
    });
    sheet.getRange(3, 1, rows.length, cols).setValues(rows);
  }

  sheet.setFrozenRows(2);
  sheet.autoResizeColumns(1, cols);
  return sheet;
}

/**
 * Builds the report for one audience ("whole" | "attended" | "rsvp").
 * Returns { audience, count, peopleScanned, sheetName, sheetUrl, message } —
 * the message is what the menu alert and the Bulk Emailer status line show.
 */
function runNoEmailReport(audience) {
  var aud = audience || BULK_AUDIENCES.WHOLE;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var [contactSheet] = sheetsByName();

  var found = collectNoEmailPeople_(contactSheet, aud);
  var sheet = writeNoEmailReportSheet_(ss, found.people, aud, found.peopleScanned);
  var audienceLabel = BULK_AUDIENCE_LABELS[aud] || aud;

  var message = (found.people.length === 0)
    ? "Zero — all " + found.peopleScanned + " people in the " + audienceLabel + " have a valid email address."
    : found.people.length + " of " + found.peopleScanned + " people in the " + audienceLabel +
      " have no usable email address. See the “" + NO_EMAIL_REPORT.SHEET_NAME + "” sheet for phone numbers and platforms.";

  Logger.log("No-Email Report [%s]: %s", aud, message);

  return {
    audience: aud,
    count: found.people.length,
    peopleScanned: found.peopleScanned,
    sheetName: NO_EMAIL_REPORT.SHEET_NAME,
    sheetUrl: ss.getUrl() + "#gid=" + sheet.getSheetId(),
    message: message
  };
}

/** Menu entry (Custom Actions → "No-Email Report…"): whole list. */
function showNoEmailReport() {
  var result = runNoEmailReport(BULK_AUDIENCES.WHOLE);
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NO_EMAIL_REPORT.SHEET_NAME)
  );
  SpreadsheetApp.getUi().alert("No-Email Report", result.message, SpreadsheetApp.getUi().ButtonSet.OK);
}
