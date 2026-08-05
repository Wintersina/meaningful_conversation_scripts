/**
 * Email Composer dialog — a popup (Custom Actions → "Email Composer…") that
 * preloads every event from the Contact List (upcoming first), shows how many
 * people each template would reach, and sends the chosen lifecycle template
 * to the people in that event's column.
 *
 * Server side of scripts/email_composer_dialog.html. Sending goes through
 * runLifecycleEmailer_ with the picked event as an explicit override, so all
 * MODE semantics and idempotency tracking behave exactly like the menu runs.
 */

function buildComposerHtml_() {
  // Inject the data at render time instead of fetching it via google.script.run:
  // dialog->server calls silently bind to the browser's DEFAULT Google session,
  // which fails with PERMISSION_DENIED when several accounts are signed in.
  // Escaping "<" keeps any "</script>"-like content from breaking the page.
  var t = HtmlService.createTemplateFromFile("email_composer_dialog");
  t.bootData = JSON.stringify(getEmailComposerData()).replace(/</g, "\\u003c");
  return t.evaluate();
}

function showEmailComposerDialog() {
  var html = buildComposerHtml_().setWidth(480).setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, "Email Composer");
}

/**
 * Web-app entry point: the SAME composer served in its own browser tab.
 * In a full tab google.script.run binds to the account that loaded the page
 * (switchable via the account chooser / ?authuser=), so Send works even in a
 * browser signed into several Google accounts — unlike the in-sheet popup.
 */
function doGet() {
  return buildComposerHtml_()
    .setTitle("Email Composer")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

// Stable /exec URL of the web-app deployment (updated by `clasp deploy -i …`).
var COMPOSER_WEBAPP_URL = "https://script.google.com/macros/s/AKfycbxzhOFuUC6a58bOZgN5qw60jIEdcVOep8fFSzXO2SttE9Qu_vUqadNvgW9J4c9HzCGI/exec";

/** Small dialog with a link that opens the composer web app in a new tab. */
function showComposerWebAppLink() {
  var url = COMPOSER_WEBAPP_URL + "?authuser=st.louis@meaningful-conversations.org";
  var html = HtmlService.createHtmlOutput(
    '<div style="font-family:Roboto,Arial,sans-serif;font-size:13px;padding:6px">' +
    '<p>Opens the Email Composer in its own browser tab, where sending works even ' +
    'when several Google accounts are signed in. If prompted, choose ' +
    '<b>st.louis@meaningful-conversations.org</b>.</p>' +
    '<p style="text-align:center"><a href="' + url + '" target="_blank" rel="noopener" ' +
    'style="display:inline-block;background:#1a73e8;color:#fff;padding:9px 18px;' +
    'border-radius:8px;text-decoration:none;font-weight:500">Open Email Composer</a></p>' +
    '</div>'
  ).setWidth(380).setHeight(170);
  SpreadsheetApp.getUi().showModalDialog(html, "Email Composer (browser tab)");
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
 * Sends from the dialog. payload = { templateKey, eventKey, mode }.
 * Returns a human-readable summary string shown in the dialog.
 */
function sendComposerEmail(payload) {
  if (!payload || !payload.templateKey || !payload.eventKey || !payload.mode) {
    throw new Error("Missing template, event, or mode.");
  }

  var config = lifecycleEmailerConfig_();
  config.MODE = payload.mode;

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
    payload.mode + ": " + parts.join(", ") + ".";
}
