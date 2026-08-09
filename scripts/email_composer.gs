/**
 * Email Composer — a popup (Custom Actions → "Email Composer…") that preloads
 * every event from the Contact List (upcoming first), shows how many people
 * each template would reach, and sends the chosen lifecycle template to the
 * people in that event's column.
 *
 * Multi-account note: the popup iframe's google.script.run calls bind to the
 * browser's DEFAULT Google session and fail with PERMISSION_DENIED when other
 * accounts are signed in. So the popup's Send instead hops through the web-app
 * deployment: it opens COMPOSER_WEBAPP_URL?action=send&…&authuser=<team acct>
 * in a small tab. doGet returns a lightweight "Sending…" page immediately;
 * that page kicks off the send via google.script.run (session-safe in a full
 * tab) and polls a CacheService progress snapshot so the user sees
 * "Sending X of Y" instead of a blank spinner. All picking still happens in
 * the popup; data is injected at render time so loading never round-trips.
 *
 * The person clicking Send types their name in the popup; it becomes both the
 * "My name is …" intro and the "Warmly, …" sign-off for that send.
 *
 * Sending goes through runLifecycleEmailer_ with the picked event as an
 * explicit override, so MODE semantics and idempotency match the menu runs.
 */

// Stable /exec URL of the web-app deployment (redeployed by the Stop hook)
// and the team account sends are pinned to.
var COMPOSER_WEBAPP_URL = "https://script.google.com/macros/s/AKfycbxzhOFuUC6a58bOZgN5qw60jIEdcVOep8fFSzXO2SttE9Qu_vUqadNvgW9J4c9HzCGI/exec";
var COMPOSER_SEND_ACCOUNT = "st.louis@meaningful-conversations.org";

/**
 * uiMode: "dialog" (in-sheet popup — Send hops via the web app URL) or
 * "webapp" (full tab — google.script.run is session-safe there, call direct).
 */
function buildComposerHtml_(uiMode) {
  // Escaping "<" keeps any "</script>"-like content from breaking the page.
  var t = HtmlService.createTemplateFromFile("email_composer_dialog");
  var data = getEmailComposerData();
  data.webappUrl = COMPOSER_WEBAPP_URL;
  data.sendAccount = COMPOSER_SEND_ACCOUNT;
  t.bootData = JSON.stringify(data).replace(/</g, "\\u003c");
  t.uiMode = uiMode;
  return t.evaluate();
}

function showEmailComposerDialog() {
  var html = buildComposerHtml_("dialog").setWidth(480).setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, "Email Composer");
}

/**
 * Web-app entry point. Without params: the composer UI in a full tab.
 * With action=send (the popup's Send hop): perform the send as the accessing
 * (authuser-pinned) account and render a small result page.
 * With page=facebook: the Facebook CSV importer UI in a full tab.
 */
function doGet(e) {
  var p = (e && e.parameter) || {};
  if (p.action === "send") {
    return buildComposerSendPage_(p);
  }
  if (p.page === "facebook") {
    return buildFacebookCsvImportHtml_("webapp")
      .setTitle("Facebook CSV Import")
      .addMetaTag("viewport", "width=device-width, initial-scale=1");
  }
  return buildComposerHtml_("webapp")
    .setTitle("Email Composer")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

/**
 * The hop tab's "Sending…" page. Returned instantly by doGet so the tab never
 * sits on a blank browser spinner; the page itself starts the send with
 * google.script.run and polls getComposerSendProgress for live counts.
 */
function buildComposerSendPage_(p) {
  var t = HtmlService.createTemplateFromFile("email_composer_send_page");
  t.payloadJson = JSON.stringify({
    templateKey: String(p.templateKey || ""),
    eventKey: String(p.eventKey || ""),
    mode: String(p.mode || ""),
    senderName: String(p.senderName || ""),
    expected: parseInt(p.expected, 10) || 0,
    progressToken: Utilities.getUuid(),
    account: Session.getActiveUser().getEmail() || "(unknown account)"
  }).replace(/</g, "\\u003c");
  return t.evaluate()
    .setTitle("Email Composer — sending")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

/**
 * Polled by the sending page (and the webapp-mode composer) for live counts.
 * Returns the snapshot written by lifecycleProgressUpdate_, or null.
 */
function getComposerSendProgress(token) {
  if (!token) return null;
  var raw = CacheService.getScriptCache().get("composerProgress:" + String(token));
  return raw ? JSON.parse(raw) : null;
}

/**
 * Data preloaded into the dialog: every event (soonest upcoming first, then
 * past, most recent first) with per-template recipient counts, plus config
 * defaults the UI displays.
 */
function getEmailComposerData() {
  var config = lifecycleEmailerConfig_();
  var [contactSheet] = sheetsByName();
  var all = getAllEventColumns_(contactSheet, config);
  var today = startOfToday_();

  // One pass over the data rows to count each event's audiences (unique valid
  // emails, mirroring collectEventAudience_'s rules).
  var lastRow = contactSheet.getLastRow();
  var data = (lastRow >= ROW_NUMBERS.ROW_12)
    ? contactSheet.getRange(ROW_NUMBERS.ROW_12, 1, lastRow - ROW_NUMBERS.ROW_12 + 1, contactSheet.getLastColumn()).getValues()
    : [];
  var emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

  var counters = all.map(function() {
    return {
      newcomers: new Set(), // welcome audience (signups minus repeat attendees)
      signups: new Set(),   // reminder audience
      noShows: new Set(),   // missed_you audience
      attended: new Set()   // follow_up audience
    };
  });

  // Welcome skips repeat attendees (see WELCOME_MAX_PRIOR_ATTENDED).
  var attendedCol0 = newcomerCapColumn_(contactSheet, EMAIL_TEMPLATES.welcome, config);

  for (var r = 0; r < data.length; r++) {
    var rawEmail = (data[r][COLUMN_INDEX.EMAIL] || "").toString().trim();
    if (!rawEmail) continue;
    var email = rawEmail.split(/[,;]+/)[0].trim().toLowerCase();
    if (!emailRegex.test(email)) continue;

    var isRegular = attendedCol0 !== -1 &&
      isRegularAttendee_(data[r][attendedCol0], config.WELCOME_MAX_PRIOR_ATTENDED);

    for (var i = 0; i < all.length; i++) {
      var cell = data[r][all[i].col0];
      if (rsvpCellIsSignup_(cell, config.INCLUDE_MAYBE_RSVPS)) {
        counters[i].signups.add(email);
        if (!isRegular) counters[i].newcomers.add(email);
      }
      if (cellIsNoShow_(cell)) counters[i].noShows.add(email);
      if (cellIsAttended_(cell)) counters[i].attended.add(email);
    }
  }

  var upcoming = [];
  var past = [];
  for (var i = 0; i < all.length; i++) {
    var e = all[i];
    var entry = {
      eventKey: e.eventKey,
      title: e.title,
      dateStr: e.dateStr,
      dayOfWeek: e.dayOfWeek,
      upcoming: e.date >= today,
      counts: {
        welcome: counters[i].newcomers.size,
        reminder: counters[i].signups.size,
        missed_you: counters[i].noShows.size,
        follow_up: counters[i].attended.size
      }
    };
    (entry.upcoming ? upcoming : past).push(entry);
  }
  past.reverse(); // most recent past event first

  return {
    upcoming: upcoming,
    past: past,
    templates: Object.keys(EMAIL_TEMPLATES).map(function(key) {
      return { key: key, label: EMAIL_TEMPLATES[key].label };
    }),
    defaultMode: config.MODE,
    testRecipient: config.TEST_RECIPIENT,
    senderName: config.SENDER_NAME
  };
}

/**
 * Sends from the dialog. payload = { templateKey, eventKey, mode, senderName,
 * progressToken }. senderName replaces both the "My name is …" intro and the
 * "Warmly, …" sign-off; progressToken enables live progress polling.
 * Returns a human-readable summary string shown in the dialog.
 */
function sendComposerEmail(payload) {
  if (!payload || !payload.templateKey || !payload.eventKey || !payload.mode) {
    throw new Error("Missing template, event, or mode.");
  }

  var config = lifecycleEmailerConfig_();
  config.MODE = payload.mode;

  var senderName = payload.senderName ? String(payload.senderName).trim() : "";
  if (senderName) {
    config.SENDER_NAME = senderName;
    config.SIGNOFF_NAME = senderName;
  }
  if (payload.progressToken) {
    config.PROGRESS_TOKEN = String(payload.progressToken);
  }

  var [contactSheet] = sheetsByName();
  var all = getAllEventColumns_(contactSheet, config);
  var picked = all.filter(function(e) { return e.eventKey === payload.eventKey; });
  if (picked.length === 0) {
    throw new Error("Event not found (was a column changed?): " + payload.eventKey);
  }

  var totals = runLifecycleEmailer_(payload.templateKey, config, null, picked);
  var tpl = EMAIL_TEMPLATES[payload.templateKey];

  var parts = [];
  if (payload.mode === "dry") parts.push(totals.planned + " would be sent");
  else parts.push(totals.sent + " sent");
  if (totals.skipped) parts.push(totals.skipped + " skipped (already sent)");
  if (totals.failed) parts.push(totals.failed + " failed");

  return tpl.label + ' → "' + picked[0].title + '" (' + picked[0].dateStr + "), mode " +
    payload.mode + ": " + parts.join(", ") + "." +
    (senderName ? ' Signed as "' + senderName + '".' : "");
}
