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

function showEmailComposerDialog() {
  // Inject the data at render time instead of fetching it via google.script.run:
  // dialog->server calls silently bind to the browser's DEFAULT Google session,
  // which fails with PERMISSION_DENIED when several accounts are signed in.
  // Escaping "<" keeps any "</script>"-like content from breaking the page.
  var t = HtmlService.createTemplateFromFile("email_composer_dialog");
  t.bootData = JSON.stringify(getEmailComposerData()).replace(/</g, "\\u003c");
  var html = t.evaluate().setWidth(480).setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, "Email Composer");
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
      signups: new Set(),   // welcome + reminder audience
      noShows: new Set(),   // missed_you audience
      attended: new Set()   // follow_up audience
    };
  });

  for (var r = 0; r < data.length; r++) {
    var rawEmail = (data[r][COLUMN_INDEX.EMAIL] || "").toString().trim();
    if (!rawEmail) continue;
    var email = rawEmail.split(/[,;]+/)[0].trim().toLowerCase();
    if (!emailRegex.test(email)) continue;

    for (var i = 0; i < all.length; i++) {
      var cell = data[r][all[i].col0];
      if (rsvpCellIsSignup_(cell, config.INCLUDE_MAYBE_RSVPS)) counters[i].signups.add(email);
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
        welcome: counters[i].signups.size,
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
 * Prompt-based composer — same picks as the popup (event, template, mode) but
 * via native Sheets prompts, which run entirely server-side. Use this when the
 * browser is signed into multiple Google accounts: google.script.run calls
 * from the HTML dialog bind to the DEFAULT session and can fail with
 * PERMISSION_DENIED, while this flow cannot.
 */
function emailComposerPromptFlow() {
  var ui = SpreadsheetApp.getUi();
  var data = getEmailComposerData();
  var events = data.upcoming.concat(data.past);
  if (events.length === 0) {
    ui.alert("Email Composer", "No events found.", ui.ButtonSet.OK);
    return;
  }

  // 1) Event
  var eventLines = events.map(function(e, i) {
    return (i + 1) + ") " + (e.upcoming ? "[upcoming] " : "[past] ") + e.title + " — " + e.dateStr;
  });
  var evResp = ui.prompt("Email Composer — 1/3: Event",
    "Enter the number of the event:\n\n" + eventLines.join("\n"), ui.ButtonSet.OK_CANCEL);
  if (evResp.getSelectedButton() !== ui.Button.OK) return;
  var evIdx = parseInt(evResp.getResponseText().trim(), 10) - 1;
  if (!(evIdx >= 0 && evIdx < events.length)) {
    ui.alert("Invalid event number."); return;
  }
  var ev = events[evIdx];

  // 2) Template
  var tplKeys = data.templates.map(function(t) { return t.key; });
  var tplLines = data.templates.map(function(t, i) {
    return (i + 1) + ") " + t.label + " — " + (ev.counts[t.key] || 0) + " people";
  });
  var tplResp = ui.prompt("Email Composer — 2/3: Email type",
    'For "' + ev.title + '" (' + ev.dateStr + "), enter the number of the email type:\n\n" + tplLines.join("\n"),
    ui.ButtonSet.OK_CANCEL);
  if (tplResp.getSelectedButton() !== ui.Button.OK) return;
  var tplIdx = parseInt(tplResp.getResponseText().trim(), 10) - 1;
  if (!(tplIdx >= 0 && tplIdx < tplKeys.length)) {
    ui.alert("Invalid email type number."); return;
  }
  var templateKey = tplKeys[tplIdx];
  var audienceCount = ev.counts[templateKey] || 0;
  if (audienceCount === 0) {
    ui.alert("No recipients match that email type for this event."); return;
  }

  // 3) Mode
  var modeResp = ui.prompt("Email Composer — 3/3: Mode",
    "Enter the mode:\n\n1) dry — log only, nothing sent\n2) test — send to " + data.testRecipient +
    " with a real individual's details\n3) actual — send REAL emails to up to " + audienceCount + " people",
    ui.ButtonSet.OK_CANCEL);
  if (modeResp.getSelectedButton() !== ui.Button.OK) return;
  var mode = ({ "1": "dry", "2": "test", "3": "actual" })[modeResp.getResponseText().trim()];
  if (!mode) { ui.alert("Invalid mode."); return; }

  if (mode === "actual") {
    var sure = ui.alert("Confirm actual send",
      'Send REAL emails to up to ' + audienceCount + ' people for "' + ev.title + '"?', ui.ButtonSet.YES_NO);
    if (sure !== ui.Button.YES) return;
  }

  var summary = sendComposerEmail({ templateKey: templateKey, eventKey: ev.eventKey, mode: mode });
  ui.alert("Email Composer", summary, ui.ButtonSet.OK);
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
