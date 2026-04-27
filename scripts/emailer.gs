function sendEmails() {
  // ——— CONFIG ————————————————————————————————————————————————————————
  var CONFIG = {
    MODE: "dry", // "dry" | "test" | "actual"
    SUBJECT: "To all Meaningful Conversations Participants:",

    // Message body source (Google Doc)
    DOC_ID: EMAILER_KEYS.docId, // required

    // Optional PDF attachment
    ATTACH_PDF: false,
    PDF_ID: EMAILER_KEYS.pdfId,

    // Test recipients for "test" mode
    TEST_RECIPIENTS: ["wintersina@gmail.com"], //"brainlift@gmail.com

    // Sheets
    CONTACT_SHEET_NAME: "Contact List",
    TRACKING_SHEET_NAME: "Email Tracking",

    // Columns in Contact List (0-based)
    COL_NAME: 0, // A: Full name
    COL_EMAIL: 5, // F: Email

    // If true, only send to rows where column E == "repeat attendee" (case-insensitive)
    FILTER_REPEAT_ATTENDEES: false,
    REPEAT_FLAG_COL: 4, // E
    REPEAT_FLAG_VALUE: "repeat attendee",

    // If true, only send to people who have an RSVP/attendance value in the
    // event column matching FILTER_EVENT_TITLE. Matches event title in Row 7
    // (normalized, case-insensitive). Cells with "-", "--", or empty are skipped.
    FILTER_BY_EVENT: true,
    FILTER_EVENT_TITLE: "One God, Many Paths",  // e.g. "A Divine Connection to Nature"

    // If true, skip emails that were already sent with the same SUBJECT in a previous run.
    // Different subjects are treated as separate sends (idempotent per email+subject).
    SKIP_ALREADY_SENT: true
  };
  // Note: Requires COL_CONSTANTS.EMAIL_START and COL_CONSTANTS.STOP_EMAIL markers
  // placed in the Name column (CONFIG.COL_NAME). Processing will start AFTER the
  // EMAIL_START row and stop BEFORE the STOP_EMAIL row.
  // ————————————————————————————————————————————————————————————————————

  // 1) Setup sheets
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var contactSheet = ss.getSheetByName(CONFIG.CONTACT_SHEET_NAME);
  if (!contactSheet) throw new Error('Missing sheet: "' + CONFIG.CONTACT_SHEET_NAME + '"');
  var tracking = ensureTrackingSheet_(ss, CONFIG.TRACKING_SHEET_NAME);

  // 2) Load message + optional PDF
  var message = loadMessageFromDoc_(CONFIG.DOC_ID);
  var attach = CONFIG.ATTACH_PDF ? loadOptionalPdf_(CONFIG.PDF_ID) : null; // {blob, name} or null

  // 3) Build recipient map (email => firstName), honoring EMAIL_START/STOP_EMAIL markers
  var recipients = (CONFIG.MODE === "test")
    ? buildTestRecipients_(CONFIG.TEST_RECIPIENTS)
    : buildUniqueRecipientsFromSheet_(
        contactSheet,
        CONFIG.COL_NAME,
        CONFIG.COLL_EMAIL, // <-- typo prevention; corrected below
        CONFIG.FILTER_REPEAT_ATTENDEES,
        CONFIG.REPEAT_FLAG_COL,
        CONFIG.REPEAT_FLAG_VALUE
      );

  // Fix minor typo: use COL_EMAIL
  // (Leaving the above call intact but reassigning here keeps the minimal-change spirit.)
  if (CONFIG.MODE !== "test") {
    // Resolve event column index if filtering by event
    var eventColIdx = -1;
    if (CONFIG.FILTER_BY_EVENT && CONFIG.FILTER_EVENT_TITLE) {
      eventColIdx = findEventColumnByTitle_(contactSheet, CONFIG.FILTER_EVENT_TITLE);
      if (eventColIdx === -1) {
        throw new Error('Event title not found in Row 7: "' + CONFIG.FILTER_EVENT_TITLE + '"');
      }
      Logger.log("Filtering by event: \"" + CONFIG.FILTER_EVENT_TITLE + "\" (column " + columnToLetter(eventColIdx + 1) + ")");
    }

    recipients = buildUniqueRecipientsFromSheet_(
      contactSheet,
      CONFIG.COL_NAME,
      CONFIG.COL_EMAIL,
      CONFIG.FILTER_REPEAT_ATTENDEES,
      CONFIG.REPEAT_FLAG_COL,
      CONFIG.REPEAT_FLAG_VALUE,
      eventColIdx
    );
  }

  // 4) Determine previously sent emails for this subject (optionally skip)
  var alreadySent = CONFIG.SKIP_ALREADY_SENT ? buildSentSet_(tracking, CONFIG.SUBJECT) : new Set();

  // 5) Dispatch per-mode
  if (CONFIG.MODE === "dry") {
    dryRunFlow_(recipients, tracking, message, attach, CONFIG.SUBJECT, alreadySent);
  } else if (CONFIG.MODE === "test") {
    testRunFlow_(recipients, tracking, message, attach, CONFIG.SUBJECT);
  } else if (CONFIG.MODE === "actual") {
    actualRunFlow_(recipients, tracking, message, attach, CONFIG.SUBJECT, alreadySent);
  } else {
    throw new Error('Unknown MODE "' + CONFIG.MODE + '" (use "dry" | "test" | "actual")');
  }

  Logger.log("Emails processed. Mode=" + CONFIG.MODE + ", attach_pdf=" + !!attach);
}

/** ————————————————————————————————————————————————————————
 * Helpers: sheets, loading, recipients, tracking
 * ———————————————————————————————————————————————————————— */
function ensureTrackingSheet_(ss, name) {
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(["Email", "Sent Status", "Run Type", "Name", "Timestamp", "Error", "Attachment", "Subject"]);
  } else {
    // Ensure header row exists and includes "Subject" in col H
    var lastCol = Math.max(sh.getLastColumn(), 8);
    var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0] || [];
    if (!headers.length || headers[0] !== "Email") {
      // Recreate full header if first row isn't a header row
      sh.insertRows(1, 1);
      sh.getRange(1, 1, 1, 8)
        .setValues([["Email", "Sent Status", "Run Type", "Name", "Timestamp", "Error", "Attachment", "Subject"]]);
    } else if (headers.length < 8 || headers[7] !== "Subject") {
      // Add/ensure Subject header in column H
      sh.getRange(1, 8).setValue("Subject");
    }
  }
  return sh;
}

function loadMessageFromDoc_(docId) {
  try {
    var doc = DocumentApp.openById(docId);
    return doc.getBody().getText();
  } catch (e) {
    throw new Error("Cannot access the message doc: " + e);
  }
}

function loadOptionalPdf_(pdfId) {
  if (!pdfId) return null;
  try {
    var file = DriveApp.getFileById(pdfId);
    return { blob: file.getBlob(), name: file.getName() };
  } catch (e) {
    // If ATTACH_PDF=true but file fetch fails, we fail fast so it isn't silent
    throw new Error("Cannot access the PDF file: " + e);
  }
}

function buildTestRecipients_(emails) {
  var emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  var map = new Map();
  for (var i = 0; i < emails.length; i++) {
    var em = (emails[i] || "").toString().trim();
    if (emailRegex.test(em) && !map.has(em)) map.set(em, "Test");
  }
  return map;
}

/**
 * Finds the 0-based column index in the Contact List whose Row 7 title
 * matches the given event title (case-insensitive, normalized).
 * Returns -1 if not found.
 */
function findEventColumnByTitle_(sheet, title) {
  var lastCol = sheet.getLastColumn();
  var startCol = HELPER_CONSTANTS.EVENT_NAMES_START_COL; // O = 15
  if (lastCol < startCol) return -1;

  var titles = sheet.getRange(ROW_NUMBERS.ROW_7, 1, 1, lastCol).getValues()[0];
  var normTarget = normalizeByStrippingWhiteSpaceAtTheEnd(title);

  for (var i = startCol - 1; i < titles.length; i++) { // 0-based
    var normTitle = normalizeByStrippingWhiteSpaceAtTheEnd(titles[i]);
    if (normTitle && normTitle === normTarget) {
      return i; // 0-based column index
    }
  }

  return -1;
}

/**
 * Build recipients honoring two markers in the Name column:
 *  - COL_CONSTANTS.EMAIL_START: start sending AFTER this row
 *  - COL_CONSTANTS.STOP_EMAIL : stop sending BEFORE this row
 *
 * If EMAIL_START appears after STOP_EMAIL, abort the run.
 *
 * Optional filters:
 *  - If filterRepeat is true, only include rows where data[r][repeatColIdx] equals repeatValue
 *    (case-insensitive, trimmed).
 *  - If eventColIdx >= 0, only include rows that have a non-empty, non-dash value
 *    in that column (people who registered/attended for that event).
 */
function buildUniqueRecipientsFromSheet_(sheet, nameColIdx, emailColIdx, filterRepeat, repeatColIdx, repeatValue, eventColIdx) {
  var data = sheet.getDataRange().getValues();
  var emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

  // Find marker rows (0-based indices into `data`)
  var startMarkerRow = -1;
  var stopMarkerRow  = -1;

  for (var i = 1; i < data.length; i++) { // skip header row
    var nameCell = (data[i][nameColIdx] || "").toString().trim();
    if (startMarkerRow === -1 && nameCell === COL_CONSTANTS.EMAIL_START) {
      startMarkerRow = i;
    }
    if (stopMarkerRow === -1 && nameCell === COL_CONSTANTS.STOP_EMAIL) {
      stopMarkerRow = i;
    }
    // Keep scanning to catch both markers even if one appears first
  }

  // Determine inclusive working bounds in terms of real data rows
  // Start AFTER the EMAIL_START row; End BEFORE the STOP_EMAIL row
  var startIdx = (startMarkerRow !== -1) ? startMarkerRow + 1 : 1;                 // default row 2
  var endIdx   = (stopMarkerRow  !== -1) ? stopMarkerRow  - 1 : (data.length - 1); // default last row

  // Validate order: start must come before end
  if (startMarkerRow !== -1 && stopMarkerRow !== -1 && startMarkerRow >= stopMarkerRow) {
    throw new Error("EMAIL_START appears after or on the same row as STOP_EMAIL. Start must come before end. Aborting run.");
  }

  // If range is empty or inverted (e.g., markers adjacent), abort
  if (endIdx < startIdx) {
    throw new Error("Computed email range is empty or invalid (start row > end row). Aborting run.");
  }

  // Log the chosen range in 1-based sheet coordinates for clarity
  Logger.log("Email range: rows %s to %s (inclusive).", (startIdx + 1), (endIdx + 1));

  // Prepare filter comparator (only if enabled)
  var shouldFilter = !!filterRepeat;
  var repeatTarget = (repeatValue || "").toString().trim().toLowerCase();

  // Event column filter
  var shouldFilterEvent = (eventColIdx !== undefined && eventColIdx >= 0);

  // Collect unique recipients within [startIdx, endIdx]
  var map = new Map();
  for (var r = startIdx; r <= endIdx; r++) {
    var row = data[r];

    // Optional repeat-attendee filter
    if (shouldFilter) {
      var cellVal = (row[repeatColIdx] || "").toString().trim().toLowerCase();
      if (cellVal !== repeatTarget) continue;
    }

    // Optional event column filter: skip rows with no RSVP/attendance for this event
    if (shouldFilterEvent) {
      var eventVal = (row[eventColIdx] || "").toString().trim();
      if (!eventVal || eventVal === "-" || eventVal === "--") continue;
    }

    var nameCell = (row[nameColIdx] || "").toString().trim();
    var email = (row[emailColIdx] || "").toString().trim();

    // Handle multi-email cells (comma-separated) — add each valid email
    var emails = email.split(/[,;]+/);
    for (var e = 0; e < emails.length; e++) {
      var singleEmail = emails[e].trim();
      if (singleEmail && emailRegex.test(singleEmail) && !map.has(singleEmail)) {
        var first = nameCell ? nameCell.split(/\s+/)[0] : "";
        map.set(singleEmail, first);
      }
    }
  }

  if (shouldFilterEvent) {
    Logger.log("Event filter: " + map.size + " recipients with RSVP/attendance in event column");
  }

  return map;
}

/**
 * Builds a Set of emails that were already sent with a specific subject.
 * Matches on email (col A) + status "Sent" (col B) + subject (col H).
 * Case-insensitive subject comparison for safety.
 */
function buildSentSet_(trackingSheet, subject) {
  var vals = trackingSheet.getDataRange().getValues();
  var sent = new Set();
  var normSubject = (subject || "").toString().trim().toLowerCase();

  for (var r = 1; r < vals.length; r++) {
    var email = (vals[r][0] || "").toString().trim();
    var status = (vals[r][1] || "").toString().trim();
    var rowSubject = (vals[r][7] || "").toString().trim().toLowerCase(); // col H = subject

    if (email && status === "Sent" && rowSubject === normSubject) {
      sent.add(email);
    }
  }
  return sent;
}

function appendTracking_(trackingSheet, email, status, runType, firstName, err, attachmentLabel, subject) {
  trackingSheet.appendRow([
    email,
    status,                // "Sent" | "Failed" | "Pending"
    runType,               // "Dry Run" | "Test Run" | "Actual Run"
    firstName,
    new Date(),
    err || "",
    attachmentLabel || "", // e.g., pdf name or "None"
    subject || ""          // Column H: Subject
  ]);
}


/** ————————————————————————————————————————————————————————
 * Core email send wrapper (handles optional attachment + quota)
 * ———————————————————————————————————————————————————————— */
function safeSendEmail_(email, subject, body, attachObj) {
  if (typeof MailApp.getRemainingDailyQuota === "function") {
    var quota = MailApp.getRemainingDailyQuota();
    if (quota <= 0) {
      return { ok: false, error: "Rate limit reached before send" };
    }
  }
  var options = {};
  if (attachObj && attachObj.blob) options.attachments = [attachObj.blob];

  try {
    MailApp.sendEmail(email, subject, body, options);
    return { ok: true, error: null };
  } catch (e) {
    return { ok: false, error: (e && e.message) ? e.message : "Unknown error" };
  }
}

/** ————————————————————————————————————————————————————————
 * Test, dryrun and Actual flows
 * ———————————————————————————————————————————————————————— */
function dryRunFlow_(recipients, tracking, message, attachObj, subject, alreadySent) {
  recipients.forEach(function(firstName, email) {
    if (alreadySent.has(email)) return; // mirror actual behavior
    var body = message;
    Logger.log("Dry run: Would send to %s (%s) attachment: %s", firstName, email, attachObj ? attachObj.name : "None");
    appendTracking_(tracking, email, "Pending", "Dry Run", firstName, "", attachObj ? attachObj.name : "None", subject);
  });
}

function testRunFlow_(recipients, tracking, message, attachObj, subject) {
  recipients.forEach(function(firstName, email) {
    var body = message;
    var res = safeSendEmail_(email, subject, body, attachObj);
    if (res.ok) {
      Logger.log("Test email sent to: %s (%s) with attachment: %s", firstName, email, attachObj ? attachObj.name : "None");
      appendTracking_(tracking, email, "Sent", "Test Run", firstName, "", attachObj ? attachObj.name : "None", subject);
    } else {
      Logger.log("Failed test send to: %s. Error: %s", email, res.error);
      appendTracking_(tracking, email, "Failed", "Test Run", firstName, res.error, attachObj ? attachObj.name : "None", subject);
    }
  });
}

function actualRunFlow_(recipients, tracking, message, attachObj, subject, alreadySent) {
  // Quick pre-check: if quota is 0, mark all as failed (not already sent)
  if (typeof MailApp.getRemainingDailyQuota === "function" && MailApp.getRemainingDailyQuota() <= 0) {
    recipients.forEach(function(firstName, email) {
      if (!alreadySent.has(email)) {
       appendTracking_(tracking, email, "Failed", "Actual Run", firstName, "Rate limit reached before send", attachObj ? attachObj.name : "None", subject);
      }
    });
    Logger.log("Aborting: Rate limit reached.");
    return;
  }

  recipients.forEach(function(firstName, email) {
    if (alreadySent.has(email)) return;
    var body = message;
    var res = safeSendEmail_(email, subject, body, attachObj);
    if (res.ok) {
      Logger.log("Email sent to: %s (%s) with attachment: %s", firstName, email, attachObj ? attachObj.name : "None");
      appendTracking_(tracking, email, "Sent", "Actual Run", firstName, "", attachObj ? attachObj.name : "None", subject);

    } else {
      Logger.log("Failed to send email to: %s Error: %s", email, res.error);
      appendTracking_(tracking, email, "Failed", "Actual Run", firstName, res.error, attachObj ? attachObj.name : "None", subject);

    }
  });
}