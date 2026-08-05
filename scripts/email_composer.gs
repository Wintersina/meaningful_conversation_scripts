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
  var html = HtmlService.createHtmlOutputFromFile("email_composer_dialog")
    .setWidth(480)
    .setHeight(600);
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
