/**
 * Bulk Emailer — sends a typed subject + body (wrapped in a clean HTML
 * template) to a chosen slice of the Contact List:
 *
 *   whole     — every data row (Row 13 down; section copies dedup by email)
 *   attended  — the Attended section (rows 13 .. above the "RSVP 2+" marker)
 *   rsvp      — the RSVP 2+ section (between "RSVP 2+" and "Stop RSVP")
 *
 * Section bounds mirror sortAttendedRows / sortRSVPRows exactly. The old
 * "Start Email" / "Stop Email" column-A markers and the Google-Doc message
 * source are gone: everything is configured in the Bulk Emailer UI.
 *
 * UI: Custom Actions → "Bulk Emailer…" opens a small launcher popup whose
 * button opens the full web-app tab (?page=bulkemailer) pinned to the team
 * account. In that tab google.script.run is session-safe, so the form loads
 * data, uploads attachments from the browser, sends via sendBulkEmails, and
 * polls the same CacheService progress snapshots the Email Composer uses.
 *
 * Idempotency: sends are logged to the "Email Tracking" sheet keyed by
 * email + subject; "skip already sent" drops addresses that already got an
 * email with the same subject.
 */

var BULK_EMAILER_DEFAULTS = {
  SUBJECT: "A Common Endeavor St. Louis",
  TEST_RECIPIENTS: ["wintersina@gmail.com"],
  TRACKING_SHEET_NAME: "Email Tracking",

  // Columns in Contact List (0-based)
  COL_NAME: 0, // A: Full name
  COL_EMAIL: 5, // F: Email

  REPEAT_FLAG_COL: 4, // E
  REPEAT_FLAG_VALUE: "repeat attendee",

  BATCH_RECIPIENT_MODE: "bcc", // "bcc" (hidden) | "to" (everyone visible)
  BATCH_SIZE: 45 // clamped to the 50-recipients-per-message cap at send time
};

var BULK_AUDIENCES = { WHOLE: "whole", ATTENDED: "attended", RSVP: "rsvp" };

/**
 * 0-based [start, end] row bounds (inclusive, into getDataRange().getValues())
 * for an audience, delimited exactly the way the sort scripts do it:
 *  - attended: sheet rows 13 .. the row above the "RSVP 2+" marker (col B)
 *  - rsvp:     two rows below "RSVP 2+" .. two rows above "Stop RSVP" (col A)
 *  - whole:    sheet rows 13 .. last row
 */
function bulkAudienceBounds_(data, audience) {
  var firstData0 = 12; // sheet row 13, where sortAttendedRows starts

  if (audience === BULK_AUDIENCES.WHOLE) {
    return { start: firstData0, end: data.length - 1 };
  }

  var rsvpIdx0 = -1;
  var stopIdx0 = -1;
  for (var i = 0; i < data.length; i++) {
    if (rsvpIdx0 === -1 && String(data[i][1] || "").trim() === COL_CONSTANTS.RSVP_2_PLUS) rsvpIdx0 = i;
    if (stopIdx0 === -1 && String(data[i][0] || "").trim() === COL_CONSTANTS.STOP_RSVP) stopIdx0 = i;
  }

  if (audience === BULK_AUDIENCES.ATTENDED) {
    if (rsvpIdx0 === -1) throw new Error('"' + COL_CONSTANTS.RSVP_2_PLUS + '" marker not found in column B — cannot bound the Attended section.');
    return { start: firstData0, end: rsvpIdx0 - 1 };
  }

  if (audience === BULK_AUDIENCES.RSVP) {
    if (rsvpIdx0 === -1) throw new Error('"' + COL_CONSTANTS.RSVP_2_PLUS + '" marker not found in column B — cannot bound the RSVP section.');
    if (stopIdx0 === -1) throw new Error('"' + COL_CONSTANTS.STOP_RSVP + '" marker not found in column A — cannot bound the RSVP section.');
    return { start: rsvpIdx0 + 2, end: stopIdx0 - 2 };
  }

  throw new Error('Unknown audience "' + audience + '" (use whole | attended | rsvp)');
}

/**
 * Wraps the typed plain-text body in a clean, email-client-safe HTML card.
 * Text is escaped so whatever is typed can't break the markup; blank lines
 * become paragraph breaks, single newlines become line breaks.
 */
function buildBulkEmailHtml_(bodyText) {
  var paragraphs = String(bodyText || "").trim().split(/\n{2,}/).map(function(p) {
    return '<p style="margin:0 0 16px;">' + escapeHtml_(p).replace(/\n/g, "<br>") + "</p>";
  }).join("");

  return (
    '<div style="margin:0;padding:24px 12px;background-color:#f5f4f0;">' +
      '<div style="max-width:600px;margin:0 auto;background-color:#ffffff;border-radius:8px;' +
        "padding:36px 40px;font-family:Georgia,'Times New Roman',serif;font-size:16px;" +
        'line-height:1.65;color:#2b2b2b;">' +
        paragraphs +
      "</div>" +
      '<div style="max-width:600px;margin:12px auto 0;text-align:center;' +
        'font-family:Georgia,serif;font-size:12px;color:#8a8a86;">Meaningful Conversations · St. Louis, MO</div>' +
    "</div>"
  );
}

/** ————————————————————————————————————————————————————————
 * Bulk Emailer UI (launcher popup + web-app tab)
 * ———————————————————————————————————————————————————————— */

/**
 * uiMode: "dialog" (in-sheet popup — just a launcher button, because iframe
 * google.script.run binds to the browser's default session) or "webapp"
 * (full tab pinned to the team account — the real form).
 */
function buildBulkEmailerHtml_(uiMode) {
  var t = HtmlService.createTemplateFromFile("bulk_emailer_dialog");
  var data = getBulkEmailerData();
  data.webappUrl = COMPOSER_WEBAPP_URL;
  data.sendAccount = COMPOSER_SEND_ACCOUNT;
  t.bootData = JSON.stringify(data).replace(/</g, "\\u003c");
  t.uiMode = uiMode;
  return t.evaluate();
}

function showBulkEmailerDialog() {
  var html = buildBulkEmailerHtml_("dialog").setWidth(420).setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, "Bulk Emailer");
}

/**
 * Boot data for the form: per-audience unique-email counts, event titles for
 * the optional event filter, and editable defaults.
 */
function getBulkEmailerData() {
  var [contactSheet] = sheetsByName();
  var data = contactSheet.getDataRange().getValues();
  var emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

  var counts = {};
  [BULK_AUDIENCES.WHOLE, BULK_AUDIENCES.ATTENDED, BULK_AUDIENCES.RSVP].forEach(function(aud) {
    try {
      var b = bulkAudienceBounds_(data, aud);
      var set = new Set();
      for (var r = b.start; r <= b.end && r < data.length; r++) {
        String(data[r][BULK_EMAILER_DEFAULTS.COL_EMAIL] || "").split(/[,;]+/).forEach(function(e) {
          var em = e.trim();
          if (emailRegex.test(em)) set.add(em.toLowerCase());
        });
      }
      counts[aud] = set.size;
    } catch (e) {
      counts[aud] = null; // markers missing — the UI shows the section as unavailable
    }
  });

  var events = getAllEventColumns_(contactSheet, lifecycleEmailerConfig_())
    .map(function(e) { return { title: e.title, dateStr: e.dateStr }; });
  events.reverse(); // newest first

  return {
    counts: counts,
    events: events,
    defaults: {
      subject: BULK_EMAILER_DEFAULTS.SUBJECT,
      testRecipients: BULK_EMAILER_DEFAULTS.TEST_RECIPIENTS.join(", "),
      batchSize: BULK_EMAILER_DEFAULTS.BATCH_SIZE,
      batchMode: BULK_EMAILER_DEFAULTS.BATCH_RECIPIENT_MODE,
      repeatFlagValue: BULK_EMAILER_DEFAULTS.REPEAT_FLAG_VALUE
    }
  };
}

/**
 * Sends from the Bulk Emailer form. payload:
 *   {
 *     mode: "dry"|"test"|"actual", audience: "whole"|"attended"|"rsvp",
 *     subject, bodyText, testRecipients: "a@x, b@y",
 *     attachment: { name, mimeType, dataB64 } | null,
 *     filterRepeat: bool, eventTitle: "" | title,
 *     attendedMoreThan: number|null, attendedLessThan: number|null,
 *     excludeEmails: "a@x, b@y", skipAlreadySent: bool,
 *     batchSingle: bool, batchMode: "bcc"|"to", batchSize: number,
 *     progressToken: string (optional, enables live progress polling)
 *   }
 * Returns a human-readable summary string shown in the form.
 */
function sendBulkEmails(payload) {
  if (!payload) throw new Error("Missing payload.");
  if (!["dry", "test", "actual"].includes(payload.mode)) {
    throw new Error('Unknown MODE "' + payload.mode + '" (use "dry" | "test" | "actual")');
  }
  var subject = String(payload.subject || "").trim();
  var bodyText = String(payload.bodyText || "").trim();
  if (!subject) throw new Error("Subject is required.");
  if (!bodyText) throw new Error("Body is required.");

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var [contactSheet] = sheetsByName();
  var tracking = ensureTrackingSheet_(ss, BULK_EMAILER_DEFAULTS.TRACKING_SHEET_NAME);

  var message = { text: bodyText, html: buildBulkEmailHtml_(bodyText) };

  var attach = null;
  if (payload.attachment && payload.attachment.dataB64) {
    attach = {
      blob: Utilities.newBlob(
        Utilities.base64Decode(payload.attachment.dataB64),
        payload.attachment.mimeType || "application/octet-stream",
        payload.attachment.name || "attachment"
      ),
      name: payload.attachment.name || "attachment"
    };
  }

  // Build recipients
  var recipients;
  var audienceLabel;
  if (payload.mode === "test") {
    var testList = String(payload.testRecipients || "").split(/[,;\n]+/).map(function(s) { return s.trim(); }).filter(Boolean);
    if (testList.length === 0) throw new Error("Test mode needs at least one test recipient.");
    recipients = buildTestRecipients_(testList);
    audienceLabel = "test recipients";
  } else {
    var audience = payload.audience || BULK_AUDIENCES.WHOLE;
    audienceLabel = audience === BULK_AUDIENCES.WHOLE ? "whole list"
      : audience === BULK_AUDIENCES.ATTENDED ? "Attended section" : "RSVP 2+ section";

    var eventColIdx = -1;
    if (payload.eventTitle) {
      eventColIdx = findEventColumnByTitle_(contactSheet, payload.eventTitle);
      if (eventColIdx === -1) throw new Error('Event title not found in Row 7: "' + payload.eventTitle + '"');
    }

    var moreThan = (payload.attendedMoreThan == null || payload.attendedMoreThan === "") ? null : Number(payload.attendedMoreThan);
    var lessThan = (payload.attendedLessThan == null || payload.attendedLessThan === "") ? null : Number(payload.attendedLessThan);
    var attendedColIdx = -1;
    if (moreThan != null || lessThan != null) {
      attendedColIdx = findAttendedCountColumn_(contactSheet);
      if (attendedColIdx === -1) throw new Error('Column "' + COL_CONSTANTS.EVENTS_ATTENDED + '" not found in Row 5.');
    }

    var excludeEmails = String(payload.excludeEmails || "").split(/[,;\n]+/).map(function(s) { return s.trim(); }).filter(Boolean);

    recipients = buildUniqueRecipientsFromSheet_(contactSheet, {
      audience: audience,
      filterRepeat: !!payload.filterRepeat,
      eventColIdx: eventColIdx,
      attendedColIdx: attendedColIdx,
      attendedMoreThan: moreThan,
      attendedLessThan: lessThan,
      excludeEmails: excludeEmails
    });
  }

  var alreadySent = payload.skipAlreadySent ? buildSentSet_(tracking, subject) : new Set();
  // Test mode never skips already-sent (it's meant for repeated verification).
  var skipSet = (payload.mode === "test") ? new Set() : alreadySent;

  bulkProgressStart_(payload.progressToken, recipients.size);

  var totals;
  if (payload.batchSingle !== false) {
    totals = batchFlow_(recipients, tracking, message, attach, subject, skipSet, payload.mode,
      payload.batchMode || BULK_EMAILER_DEFAULTS.BATCH_RECIPIENT_MODE,
      Number(payload.batchSize) || BULK_EMAILER_DEFAULTS.BATCH_SIZE);
  } else if (payload.mode === "dry") {
    totals = dryRunFlow_(recipients, tracking, message, attach, subject, skipSet);
  } else if (payload.mode === "test") {
    totals = testRunFlow_(recipients, tracking, message, attach, subject);
  } else {
    totals = actualRunFlow_(recipients, tracking, message, attach, subject, skipSet);
  }

  bulkProgressFlush_(true);

  var parts = [];
  if (payload.mode === "dry") parts.push(totals.planned + " would be sent");
  else parts.push(totals.sent + " sent");
  if (totals.skipped) parts.push(totals.skipped + " skipped (already sent)");
  if (totals.failed) parts.push(totals.failed + " failed");
  if (totals.messages != null && payload.mode !== "dry") parts.push(totals.messages + " message(s)");

  return 'Bulk email "' + subject + '" → ' + audienceLabel + ", mode " + payload.mode + ": " +
    parts.join(", ") + "." + (attach ? ' Attachment: "' + attach.name + '".' : "");
}

/** ————————————————————————————————————————————————————————
 * Live-progress snapshots (same cache keys the Email Composer polls
 * via getComposerSendProgress)
 * ———————————————————————————————————————————————————————— */
var bulkProgress_ = null;

function bulkProgressStart_(token, total) {
  if (!token) { bulkProgress_ = null; return; }
  bulkProgress_ = { token: token, total: total, processed: 0, sent: 0, skipped: 0, failed: 0, planned: 0 };
  bulkProgressFlush_(false);
}

/** outcome: "sent" | "skipped" | "failed" | "planned"; count defaults to 1. */
function bulkProgressStep_(outcome, count) {
  if (!bulkProgress_) return;
  var n = (count == null) ? 1 : count;
  bulkProgress_.processed += n;
  if (outcome && bulkProgress_[outcome] != null) bulkProgress_[outcome] += n;
  bulkProgressFlush_(false);
}

function bulkProgressFlush_(done) {
  if (!bulkProgress_) return;
  try {
    CacheService.getScriptCache().put(
      "composerProgress:" + bulkProgress_.token,
      JSON.stringify({
        total: bulkProgress_.total,
        processed: bulkProgress_.processed,
        sent: bulkProgress_.sent,
        skipped: bulkProgress_.skipped,
        failed: bulkProgress_.failed,
        planned: bulkProgress_.planned,
        done: !!done
      }),
      600
    );
  } catch (e) {} // progress is best-effort — never break a send over it
}

/** ————————————————————————————————————————————————————————
 * Helpers: sheets, recipients, tracking
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

/**
 * Normalizes an email body into { text, html }. Accepts either a plain string
 * (html omitted) or a { text, html } object. Keeps send helpers robust
 * regardless of which form the caller passes.
 */
function asEmailContent_(content) {
  if (content && typeof content === "object") {
    return { text: content.text || "", html: content.html || null };
  }
  return { text: String(content || ""), html: null };
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
 * Finds the 0-based column index of the "# Events Attended" column by looking
 * up COL_CONSTANTS.EVENTS_ATTENDED in the Row 5 header (same pattern as
 * backfill_events_formulas / calendar_sync). Returns -1 if not found.
 */
function findAttendedCountColumn_(sheet) {
  var headerRow = sheet.getRange(ROW_NUMBERS.ROW_5, 1, 1, sheet.getLastColumn()).getValues()[0];
  return headerRow.indexOf(COL_CONSTANTS.EVENTS_ATTENDED);
}

/**
 * Builds a unique email → firstName map for an audience slice of the
 * Contact List. opts:
 *   audience         — "whole" | "attended" | "rsvp" (see bulkAudienceBounds_)
 *   filterRepeat     — only rows where col E == "repeat attendee"
 *   eventColIdx      — >= 0: only rows with a non-empty, non-dash value there
 *   attendedColIdx   — >= 0: apply the attended-count bounds below
 *   attendedMoreThan — strictly-greater bound (null = off)
 *   attendedLessThan — strictly-less bound (null = off)
 *   excludeEmails    — addresses to always drop (case-insensitive)
 */
function buildUniqueRecipientsFromSheet_(sheet, opts) {
  var data = sheet.getDataRange().getValues();
  var emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

  var bounds = bulkAudienceBounds_(data, opts.audience || BULK_AUDIENCES.WHOLE);
  var startIdx = bounds.start;
  var endIdx = Math.min(bounds.end, data.length - 1);
  if (endIdx < startIdx) {
    throw new Error("Computed audience range is empty or invalid (start row > end row). Aborting run.");
  }
  Logger.log("Audience '%s': rows %s to %s (inclusive).", opts.audience, startIdx + 1, endIdx + 1);

  var repeatTarget = BULK_EMAILER_DEFAULTS.REPEAT_FLAG_VALUE.toLowerCase();
  var shouldFilterEvent = (opts.eventColIdx != null && opts.eventColIdx >= 0);
  var shouldFilterAttended = (opts.attendedColIdx != null && opts.attendedColIdx >= 0);
  var attendedMin = (opts.attendedMoreThan != null) ? Number(opts.attendedMoreThan) : null;
  var attendedMax = (opts.attendedLessThan != null) ? Number(opts.attendedLessThan) : null;

  var excludeSet = new Set();
  (opts.excludeEmails || []).forEach(function(e) {
    var norm = (e || "").toString().trim().toLowerCase();
    if (norm) excludeSet.add(norm);
  });
  var excludedCount = 0;

  var map = new Map();
  for (var r = startIdx; r <= endIdx; r++) {
    var row = data[r];

    if (opts.filterRepeat) {
      var cellVal = (row[BULK_EMAILER_DEFAULTS.REPEAT_FLAG_COL] || "").toString().trim().toLowerCase();
      if (cellVal !== repeatTarget) continue;
    }

    if (shouldFilterEvent) {
      var eventVal = (row[opts.eventColIdx] || "").toString().trim();
      if (!eventVal || eventVal === "-" || eventVal === "--") continue;
    }

    if (shouldFilterAttended) {
      var attendedCount = Number(row[opts.attendedColIdx]);
      if (!isFinite(attendedCount)) continue;
      if (attendedMin !== null && attendedCount <= attendedMin) continue;
      if (attendedMax !== null && attendedCount >= attendedMax) continue;
    }

    var nameCell = (row[BULK_EMAILER_DEFAULTS.COL_NAME] || "").toString().trim();
    var email = (row[BULK_EMAILER_DEFAULTS.COL_EMAIL] || "").toString().trim();

    // Handle multi-email cells (comma-separated) — add each valid email
    var emails = email.split(/[,;]+/);
    for (var e = 0; e < emails.length; e++) {
      var singleEmail = emails[e].trim();
      if (!singleEmail || !emailRegex.test(singleEmail) || map.has(singleEmail)) continue;
      if (excludeSet.has(singleEmail.toLowerCase())) {
        excludedCount++;
        continue;
      }
      var first = nameCell ? nameCell.split(/\s+/)[0] : "";
      map.set(singleEmail, first);
    }
  }

  if (shouldFilterEvent) {
    Logger.log("Event filter: " + map.size + " recipients with RSVP/attendance in event column");
  }
  if (shouldFilterAttended) {
    var boundsDesc = [];
    if (attendedMin !== null) boundsDesc.push("more than " + attendedMin);
    if (attendedMax !== null) boundsDesc.push("less than " + attendedMax);
    Logger.log("Attendance filter: " + map.size + " recipients with " + boundsDesc.join(" and ") + " meetings attended");
  }
  if (excludeSet.size > 0) {
    Logger.log("Exclude list: dropped " + excludedCount + " address(es) matching the exclude list");
  }

  return map;
}

/**
 * Builds a Set of emails that were already sent with a specific subject.
 * Matches on email (col A) + status "Sent" (col B) + subject (col H).
 * Case-insensitive subject comparison for safety.
 *
 * The Email cell may hold a single address (per-recipient rows) or many addresses
 * stacked newline/comma-separated (batch rows), so each cell is split and every
 * address is registered — keeping skip-already-sent correct in both modes.
 */
function buildSentSet_(trackingSheet, subject) {
  var vals = trackingSheet.getDataRange().getValues();
  var sent = new Set();
  var normSubject = (subject || "").toString().trim().toLowerCase();

  for (var r = 1; r < vals.length; r++) {
    var emailCell = (vals[r][0] || "").toString().trim();
    var status = (vals[r][1] || "").toString().trim();
    var rowSubject = (vals[r][7] || "").toString().trim().toLowerCase(); // col H = subject

    if (emailCell && status === "Sent" && rowSubject === normSubject) {
      emailCell.split(/[\n,;]+/).forEach(function(e) {
        var trimmed = e.trim();
        if (trimmed) sent.add(trimmed);
      });
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
    attachmentLabel || "", // e.g., attachment file name or "None"
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
  var content = asEmailContent_(body);
  var options = {};
  if (content.html) options.htmlBody = content.html;
  if (attachObj && attachObj.blob) options.attachments = [attachObj.blob];

  try {
    MailApp.sendEmail(email, subject, content.text, options);
    return { ok: true, error: null };
  } catch (e) {
    return { ok: false, error: (e && e.message) ? e.message : "Unknown error" };
  }
}

/**
 * Sends ONE email to a list of recipients at once.
 * - recipientMode "bcc": To is the sending account, everyone else is BCC'd (addresses hidden).
 * - recipientMode "to" : all addresses go in the To field (everyone sees each other).
 * Quota is measured in recipients, so a batch of N counts as N (plus 1 for the To self in bcc mode).
 */
function safeSendBatchEmail_(emails, subject, body, attachObj, recipientMode) {
  var isBcc = (recipientMode !== "to"); // default to bcc for privacy
  var recipientCount = emails.length + (isBcc ? 1 : 0); // +1 for the To self in bcc mode

  if (typeof MailApp.getRemainingDailyQuota === "function") {
    var quota = MailApp.getRemainingDailyQuota();
    if (quota < recipientCount) {
      return { ok: false, error: "Rate limit: need " + recipientCount + " recipients but quota is " + quota };
    }
  }

  var content = asEmailContent_(body);
  var options = {};
  if (content.html) options.htmlBody = content.html;
  if (attachObj && attachObj.blob) options.attachments = [attachObj.blob];

  var toField;
  if (isBcc) {
    // Address the email to the sending account and BCC the whole list.
    toField = Session.getActiveUser().getEmail() || emails[0];
    options.bcc = emails.join(",");
  } else {
    toField = emails.join(",");
  }

  try {
    MailApp.sendEmail(toField, subject, content.text, options);
    return { ok: true, error: null };
  } catch (e) {
    return { ok: false, error: (e && e.message) ? e.message : "Unknown error" };
  }
}

/**
 * Batch flow: collects all (non-skipped) recipients and sends them in batched
 * emails of at most batchSize recipients each (to stay under the provider's
 * "recipients per message" limit). Honors MODE ("dry" logs only; "test"/"actual"
 * actually send). Because each send is one un-personalized email, it writes a
 * SINGLE tracking row per chunk: the Email column holds that chunk's addresses
 * stacked newline-separated, and the Name column holds the count. buildSentSet_
 * splits that cell back out, so skip-already-sent still works per-email.
 *
 * Returns { sent, skipped, failed, planned, messages }.
 */
function batchFlow_(recipients, tracking, message, attachObj, subject, alreadySent, mode, recipientMode, batchSize) {
  var totals = { sent: 0, skipped: 0, failed: 0, planned: 0, messages: 0 };
  var emails = [];
  recipients.forEach(function(firstName, email) {
    if (alreadySent && alreadySent.has(email)) { totals.skipped++; return; }
    emails.push(email);
  });
  if (totals.skipped) bulkProgressStep_("skipped", totals.skipped);

  var attachLabel = attachObj ? attachObj.name : "None";
  var runType = (mode === "dry") ? "Dry Run" : (mode === "test" ? "Test Run" : "Actual Run");
  var via = (recipientMode !== "to" ? "BCC" : "TO");

  if (emails.length === 0) {
    Logger.log("Batch %s: no recipients to send (all filtered or already sent).", mode);
    return totals;
  }

  // Apps Script hard cap: 50 recipients (to + cc + bcc combined) per message.
  // In BCC mode the To field holds 1 address (the sending self), so reserve a slot.
  var MAX_RECIPIENTS_PER_MESSAGE = 50;
  var reserved = (recipientMode !== "to") ? 1 : 0; // To-self occupies 1 slot in bcc mode
  var maxChunk = MAX_RECIPIENTS_PER_MESSAGE - reserved;

  // Effective chunk size: requested batchSize, clamped to the provider cap.
  var requested = (batchSize && batchSize > 0) ? batchSize : emails.length;
  var chunkSize = Math.min(requested, maxChunk);

  var chunks = [];
  for (var i = 0; i < emails.length; i += chunkSize) {
    chunks.push(emails.slice(i, i + chunkSize));
  }
  totals.messages = chunks.length;

  Logger.log("Batch %s: %s recipient(s) split into %s message(s) of up to %s via %s (cap %s/msg).",
    mode, emails.length, chunks.length, chunkSize, via, MAX_RECIPIENTS_PER_MESSAGE);

  chunks.forEach(function(chunk, idx) {
    // One tracking row per chunk: chunk's addresses stacked in the Email column.
    var emailCell = chunk.join("\n");
    var countLabel = chunk.length + " recipients (batch " + (idx + 1) + "/" + chunks.length + ")";

    if (mode === "dry") {
      Logger.log("Dry run (batch %s/%s): Would send ONE email to %s recipient(s) via %s. Attachment: %s",
        (idx + 1), chunks.length, chunk.length, via, attachLabel);
      appendTracking_(tracking, emailCell, "Pending", runType, countLabel, "", attachLabel, subject);
      totals.planned += chunk.length;
      bulkProgressStep_("planned", chunk.length);
      return;
    }

    // test or actual: actually send this chunk's email
    var res = safeSendBatchEmail_(chunk, subject, message, attachObj, recipientMode);
    if (res.ok) {
      Logger.log("Batch %s (%s/%s): sent ONE email to %s recipient(s) via %s. Attachment: %s",
        mode, (idx + 1), chunks.length, chunk.length, via, attachLabel);
      totals.sent += chunk.length;
      bulkProgressStep_("sent", chunk.length);
    } else {
      Logger.log("Batch %s (%s/%s): failed to send. Error: %s", mode, (idx + 1), chunks.length, res.error);
      totals.failed += chunk.length;
      bulkProgressStep_("failed", chunk.length);
    }
    appendTracking_(tracking, emailCell, res.ok ? "Sent" : "Failed", runType, countLabel,
      res.ok ? "" : res.error, attachLabel, subject);
  });

  return totals;
}

/** ————————————————————————————————————————————————————————
 * Per-recipient flows (one personalized tracking row per address)
 * Each returns { sent, skipped, failed, planned }.
 * ———————————————————————————————————————————————————————— */
function dryRunFlow_(recipients, tracking, message, attachObj, subject, alreadySent) {
  var totals = { sent: 0, skipped: 0, failed: 0, planned: 0 };
  recipients.forEach(function(firstName, email) {
    if (alreadySent.has(email)) { totals.skipped++; bulkProgressStep_("skipped"); return; }
    Logger.log("Dry run: Would send to %s (%s) attachment: %s", firstName, email, attachObj ? attachObj.name : "None");
    appendTracking_(tracking, email, "Pending", "Dry Run", firstName, "", attachObj ? attachObj.name : "None", subject);
    totals.planned++;
    bulkProgressStep_("planned");
  });
  return totals;
}

function testRunFlow_(recipients, tracking, message, attachObj, subject) {
  var totals = { sent: 0, skipped: 0, failed: 0, planned: 0 };
  recipients.forEach(function(firstName, email) {
    var res = safeSendEmail_(email, subject, message, attachObj);
    if (res.ok) {
      Logger.log("Test email sent to: %s (%s) with attachment: %s", firstName, email, attachObj ? attachObj.name : "None");
      appendTracking_(tracking, email, "Sent", "Test Run", firstName, "", attachObj ? attachObj.name : "None", subject);
      totals.sent++;
      bulkProgressStep_("sent");
    } else {
      Logger.log("Failed test send to: %s. Error: %s", email, res.error);
      appendTracking_(tracking, email, "Failed", "Test Run", firstName, res.error, attachObj ? attachObj.name : "None", subject);
      totals.failed++;
      bulkProgressStep_("failed");
    }
  });
  return totals;
}

function actualRunFlow_(recipients, tracking, message, attachObj, subject, alreadySent) {
  var totals = { sent: 0, skipped: 0, failed: 0, planned: 0 };

  // Quick pre-check: if quota is 0, mark all as failed (not already sent)
  if (typeof MailApp.getRemainingDailyQuota === "function" && MailApp.getRemainingDailyQuota() <= 0) {
    recipients.forEach(function(firstName, email) {
      if (!alreadySent.has(email)) {
        appendTracking_(tracking, email, "Failed", "Actual Run", firstName, "Rate limit reached before send", attachObj ? attachObj.name : "None", subject);
        totals.failed++;
        bulkProgressStep_("failed");
      }
    });
    Logger.log("Aborting: Rate limit reached.");
    return totals;
  }

  recipients.forEach(function(firstName, email) {
    if (alreadySent.has(email)) { totals.skipped++; bulkProgressStep_("skipped"); return; }
    var res = safeSendEmail_(email, subject, message, attachObj);
    if (res.ok) {
      Logger.log("Email sent to: %s (%s) with attachment: %s", firstName, email, attachObj ? attachObj.name : "None");
      appendTracking_(tracking, email, "Sent", "Actual Run", firstName, "", attachObj ? attachObj.name : "None", subject);
      totals.sent++;
      bulkProgressStep_("sent");
    } else {
      Logger.log("Failed to send email to: %s Error: %s", email, res.error);
      appendTracking_(tracking, email, "Failed", "Actual Run", firstName, res.error, attachObj ? attachObj.name : "None", subject);
      totals.failed++;
      bulkProgressStep_("failed");
    }
  });
  return totals;
}
