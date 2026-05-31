/**
 * Syncs events from the Contact List sheet to Google Calendar.
 *
 * Idempotent and self-consolidating:
 *  - The canonical Google Calendar event id for each date is stored in the
 *    Schedule sheet, column F (keyed by the date in column C). Sync looks up
 *    that id first, so it updates the existing event instead of guessing by
 *    time/title and spawning duplicates.
 *  - For each event date it gathers every calendar event in the evening window
 *    that belongs to the title family (the week's topic, a stray "TBD", or the
 *    Eventbrite "(Free Event)"), merges their useful info into one canonical
 *    event, and deletes the rest.
 */
function syncEventsToGoogleCalendar() {
  Logger.log("starting syncEventsToGoogleCalendar");
  runCalendarSync_({ allowCreate: true, skipOld: true });
}

/**
 * One-shot cleanup you can run from the menu to consolidate the existing
 * backlog of duplicate calendar events. Merges and deletes only — never
 * creates — and does not skip older dates.
 */
function consolidateExistingCalendarDuplicates() {
  Logger.log("starting consolidateExistingCalendarDuplicates");
  runCalendarSync_({ allowCreate: false, skipOld: false });
}

/**
 * Core sync/consolidation routine.
 * @param {{allowCreate: boolean, skipOld: boolean}} options
 */
function runCalendarSync_(options) {
  var allowCreate = options && options.allowCreate;
  var skipOld = options && options.skipOld;

  var calendar = CalendarApp.getDefaultCalendar();
  var contactListSheet = sheetsByName()[0];
  var scheduleSheet = sheetsByName()[2];

  // Find the "# Events Attended" column in row 5 to know where event columns end
  var lastCol = contactListSheet.getLastColumn();
  var row5Values = contactListSheet.getRange(ROW_NUMBERS.ROW_5, 1, 1, lastCol).getValues()[0];
  var attendedIndex0 = row5Values.indexOf(COL_CONSTANTS.EVENTS_ATTENDED);
  if (attendedIndex0 === -1) {
    Logger.log("Could not find '# Events Attended' column. Aborting sync.");
    return;
  }
  var attendedCol1 = attendedIndex0 + 1; // 1-based

  // Event columns: start at col O (15), end just before "# Events Attended"
  var startCol = HELPER_CONSTANTS.EVENT_NAMES_START_COL; // 15
  var endCol = attendedCol1 - 1;

  if (endCol < startCol) {
    Logger.log("No event columns found. Aborting sync.");
    return;
  }

  var numCols = endCol - startCol + 1;

  // Batch-read row 4 (rooms), row 6 (dates), row 7 (titles), and row 9 (RSVP counts) across event columns
  var rooms = contactListSheet.getRange(ROW_NUMBERS.ROW_4, startCol, 1, numCols).getValues()[0];
  var dates = contactListSheet.getRange(ROW_NUMBERS.ROW_6, startCol, 1, numCols).getValues()[0];
  var titles = contactListSheet.getRange(ROW_NUMBERS.ROW_7, startCol, 1, numCols).getValues()[0];
  var rsvpRaw = contactListSheet.getRange(ROW_NUMBERS.ROW_9, startCol, 1, numCols).getValues()[0];

  // Date -> { row, eventId } map from the Schedule sheet (our id store).
  var scheduleByDate = getScheduleDateMap_(scheduleSheet);
  ensureScheduleIdHeader_(scheduleSheet);

  var timeZone = "America/Chicago";
  var createdCount = 0;
  var updatedCount = 0;
  var deletedCount = 0;

  for (var i = 0; i < numCols; i++) {
    var titleValue = titles[i];
    if (!titleValue || String(titleValue).trim() === "") {
      continue;
    }

    var eventDate = parseEventDate_(dates[i]);
    if (!eventDate) {
      Logger.log("Skipping column " + (startCol + i) + ": could not parse date '" + dates[i] + "'");
      continue;
    }

    var title = String(titleValue).trim();
    var room = rooms[i] ? String(rooms[i]).trim() : "";
    var rsvpCount = parseRsvpCount_(rsvpRaw[i]);
    var calendarTitle = room ? title + " (" + room + ")" : title;

    var startTime = buildDateInTimeZone_(eventDate, 18, 30, timeZone);
    var endTime = buildDateInTimeZone_(eventDate, 20, 0, timeZone);

    var now = new Date();
    var nowStr = Utilities.formatDate(now, timeZone, "MMM d, yyyy h:mm a");
    var dateStr = Utilities.formatDate(startTime, timeZone, "MMMM d, yyyy");

    if (skipOld) {
      var oneMonthAgo = new Date(now.getFullYear(), now.getMonth() - 1, now.getDate());
      if (startTime < oneMonthAgo) {
        Logger.log("Skipping past event (> 1 month old): " + calendarTitle);
        continue;
      }
    }

    var key = dateKey_(eventDate);
    var scheduleInfo = scheduleByDate[key] || null;
    var storedEventId = scheduleInfo ? scheduleInfo.eventId : "";

    var result = consolidateDay_(calendar, {
      eventDate: eventDate,
      title: title,
      calendarTitle: calendarTitle,
      room: room,
      rsvpCount: rsvpCount,
      startTime: startTime,
      endTime: endTime,
      storedEventId: storedEventId,
      timeZone: timeZone,
      nowStr: nowStr,
      dateStr: dateStr,
      allowCreate: allowCreate
    });

    deletedCount += result.deleted;
    if (result.created) createdCount++;
    else if (result.canonical) updatedCount++;

    // Persist the canonical event id back into Schedule!F for this date.
    if (result.canonical) {
      var canonicalId = result.canonical.getId();
      if (scheduleInfo && scheduleInfo.row) {
        if (scheduleInfo.eventId !== canonicalId) {
          scheduleSheet.getRange(scheduleInfo.row, 6).setValue(canonicalId);
          scheduleInfo.eventId = canonicalId;
        }
      } else {
        Logger.log("No Schedule row found for " + dateStr + " — could not store event id.");
      }
    }
  }

  Logger.log("Calendar sync complete. Created: " + createdCount +
    ", Updated: " + updatedCount + ", Deleted duplicates: " + deletedCount);
}

/**
 * Finds the canonical calendar event for a date, merges sibling events
 * (duplicates, a stray "TBD", the Eventbrite "(Free Event)") into it, and
 * deletes the siblings. Optionally creates the event if none exists.
 *
 * @param {CalendarApp.Calendar} calendar
 * @param {Object} o - See call site for fields.
 * @return {{canonical: CalendarApp.CalendarEvent|null, created: boolean, deleted: number}}
 */
function consolidateDay_(calendar, o) {
  var deleted = 0;
  var created = false;

  // Gather every event that evening (5pm–9pm) so a minor time drift or a
  // differently-timed auto-event still gets caught.
  var windowStart = buildDateInTimeZone_(o.eventDate, 17, 0, o.timeZone);
  var windowEnd = buildDateInTimeZone_(o.eventDate, 21, 0, o.timeZone);
  var dayEvents = calendar.getEvents(windowStart, windowEnd);

  // Keep only the title family for this date:
  //   - a literal "TBD" placeholder, or
  //   - any title containing this week's topic (covers "(Rm 208)" and
  //     "(Free Event)" suffixes since both contain the base title).
  var normTitle = normalizeString(o.title).toLowerCase().trim();
  var family = dayEvents.filter(function (ev) {
    var t = normalizeString(ev.getTitle()).toLowerCase().trim();
    return t === "tbd" || (normTitle && t.indexOf(normTitle) !== -1);
  });

  // 1) Canonical = the stored-id event if it still exists.
  var canonical = null;
  if (o.storedEventId) {
    try {
      var byId = calendar.getEventById(o.storedEventId);
      if (byId) canonical = byId;
    } catch (e) {
      Logger.log("Stored event id lookup failed for " + o.dateStr + ": " + e.message);
    }
  }

  // 2) Otherwise prefer a script-managed event (has a "Last synced" line),
  //    else the earliest event in the family.
  if (!canonical && family.length) {
    var managed = family.filter(function (ev) {
      return (ev.getDescription() || "").indexOf("Last synced") !== -1;
    });
    var pool = managed.length ? managed : family;
    pool.sort(function (a, b) { return a.getStartTime() - b.getStartTime(); });
    canonical = pool[0];
  }

  // Harvest merge-worthy info from the whole family before deleting anything.
  var harvest = harvestFamilyInfo_(family);

  // 3) Create if nothing exists and we're allowed.
  if (!canonical) {
    if (!o.allowCreate) {
      return { canonical: null, created: false, deleted: 0 };
    }
    var description = buildDescription_(o.room, o.rsvpCount, harvest, o.nowStr);
    canonical = calendar.createEvent(o.calendarTitle, o.startTime, o.endTime, {
      description: description
    });
    created = true;
    Logger.log("Created event: " + o.calendarTitle + " on " + o.dateStr);
  }

  var canonicalId = canonical.getId();

  // Delete every other family member (log each). Falls back to the advanced
  // Calendar API for externally-managed events (e.g. the Eventbrite
  // "(Free Event)") that CalendarApp refuses to delete.
  for (var j = 0; j < family.length; j++) {
    if (family[j].getId() === canonicalId) continue;
    try {
      Logger.log("Deleting duplicate: '" + family[j].getTitle() + "' [" + family[j].getId() + "] on " + o.dateStr);
      deleteCalendarEventRobust_(calendar, family[j]);
      deleted++;
    } catch (e) {
      Logger.log("Could not delete '" + family[j].getTitle() + "' on " + o.dateStr + ": " + e.message);
    }
  }

  // Normalize the canonical event's title, time, and description. For a
  // freshly created event this is already correct, so only adopted events
  // need it — but it's idempotent either way.
  var desc = canonical.getDescription() || "";
  desc = updateDescriptionField_(desc, "Room", o.room || null);
  desc = updateDescriptionField_(desc, "RSVPs", o.rsvpCount !== null ? o.rsvpCount : null);
  if (harvest.confirmation) desc = updateDescriptionField_(desc, "Confirmation", harvest.confirmation);
  if (harvest.guests) desc = updateDescriptionField_(desc, "Eventbrite guests", harvest.guests);
  if (harvest.notes) desc = updateDescriptionField_(desc, "Notes", harvest.notes);
  desc = updateDescriptionField_(desc, "Last synced", o.nowStr);

  updateCanonicalEvent_(calendar, canonical, {
    calendarTitle: o.calendarTitle,
    startTime: o.startTime,
    endTime: o.endTime,
    description: desc,
    timeZone: o.timeZone,
    dateStr: o.dateStr
  });

  return { canonical: canonical, created: created, deleted: deleted };
}

/**
 * Pulls merge-worthy details out of a family of same-day events:
 * Eventbrite confirmation number + guest count from the "(Free Event)",
 * and room/notes from a "TBD" placeholder.
 * @param {CalendarApp.CalendarEvent[]} family
 * @return {{confirmation: string|null, guests: number|null, notes: string|null}}
 */
function harvestFamilyInfo_(family) {
  var confirmation = null;
  var guests = null;
  var noteParts = [];

  for (var i = 0; i < family.length; i++) {
    var ev = family[i];
    var desc = ev.getDescription() || "";
    var t = normalizeString(ev.getTitle()).toLowerCase().trim();

    var conf = desc.match(/Confirmation number:\s*(\d+)/i);
    if (conf) confirmation = conf[1];

    try {
      var g = ev.getGuestList();
      if (g && g.length) guests = g.length;
    } catch (e) { /* some events expose no guest list */ }

    if (t === "tbd") {
      var loc = "";
      try { loc = ev.getLocation() || ""; } catch (e) { loc = ""; }
      var body = desc.replace(/Last synced:.*$/m, "").trim();
      var combined = [loc, body].filter(function (s) { return s && s.trim() !== ""; })
        .join(" — ").replace(/\s*\n\s*/g, "; ").trim();
      if (combined) noteParts.push(combined);
    }
  }

  return {
    confirmation: confirmation,
    guests: guests,
    notes: noteParts.length ? noteParts.join(" | ") : null
  };
}

/**
 * Builds a fresh description for a newly created event.
 */
function buildDescription_(room, rsvpCount, harvest, nowStr) {
  var description = "Last synced: " + nowStr;
  if (room) description += "\nRoom: " + room;
  if (rsvpCount !== null) description += "\nRSVPs: " + rsvpCount;
  if (harvest.confirmation) description += "\nConfirmation: " + harvest.confirmation;
  if (harvest.guests) description += "\nEventbrite guests: " + harvest.guests;
  if (harvest.notes) description += "\nNotes: " + harvest.notes;
  return description;
}

/**
 * Reads the Schedule sheet into a { "yyyy-MM-dd": {row, eventId} } map,
 * keyed by the date in column C. Row is 1-based; eventId comes from column F.
 * @param {Sheet} scheduleSheet
 * @return {Object}
 */
function getScheduleDateMap_(scheduleSheet) {
  var map = {};
  var lastRow = scheduleSheet.getLastRow();
  if (lastRow < 2) return map;

  // Columns C(3) .. F(6): date, topic, location, cal event id
  var data = scheduleSheet.getRange(2, 3, lastRow - 1, 4).getValues();
  for (var i = 0; i < data.length; i++) {
    var d = parseEventDate_(data[i][0]);
    if (!d) continue;
    var key = dateKey_(d);
    // First row wins for a given date (schedule dates are unique anyway).
    if (!map[key]) {
      map[key] = {
        row: i + 2,
        eventId: data[i][3] ? String(data[i][3]).trim() : ""
      };
    }
  }
  return map;
}

/**
 * Labels the Schedule id column (F1) the first time, if blank.
 */
function ensureScheduleIdHeader_(scheduleSheet) {
  var headerCell = scheduleSheet.getRange(1, 6);
  if (!headerCell.getValue()) {
    headerCell.setValue("Cal Event ID");
  }
}

/**
 * Formats a date-only Date as a stable "yyyy-MM-dd" key from local components.
 * @param {Date} d
 * @return {string}
 */
function dateKey_(d) {
  return d.getFullYear() + "-" +
    ("0" + (d.getMonth() + 1)).slice(-2) + "-" +
    ("0" + d.getDate()).slice(-2);
}

/**
 * Patches an existing event via the Calendar advanced service.
 * Bypasses CalendarApp's stricter setters, which throw "Action not allowed"
 * on events created by external integrations even when the caller owns them.
 * Requires the Calendar advanced service to be enabled in the Apps Script editor.
 * @param {CalendarApp.Calendar} calendar
 * @param {CalendarApp.CalendarEvent} event
 * @param {Object} patch - Calendar v3 Event resource fragment (e.g. {summary, description})
 */
function patchCalendarEvent_(calendar, event, patch) {
  var eventId = event.getId().replace(/@google\.com$/, "");
  Calendar.Events.patch(patch, calendar.getId(), eventId);
}

/**
 * Deletes an event, falling back to the advanced Calendar service when
 * CalendarApp throws "Action not allowed" on externally-created events.
 * @param {CalendarApp.Calendar} calendar
 * @param {CalendarApp.CalendarEvent} event
 */
function deleteCalendarEventRobust_(calendar, event) {
  try {
    event.deleteEvent();
  } catch (e) {
    var eventId = event.getId().replace(/@google\.com$/, "");
    Calendar.Events.remove(calendar.getId(), eventId);
  }
}

/**
 * Sets title/time/description on the canonical event, falling back to a single
 * advanced-service patch when CalendarApp's setters are not allowed (external
 * integration events such as the Eventbrite "(Free Event)").
 * @param {CalendarApp.Calendar} calendar
 * @param {CalendarApp.CalendarEvent} event
 * @param {Object} o - {calendarTitle, startTime, endTime, description, timeZone, dateStr}
 */
function updateCanonicalEvent_(calendar, event, o) {
  var needTitle = event.getTitle() !== o.calendarTitle;
  var needTime = event.getStartTime().getTime() !== o.startTime.getTime() ||
    event.getEndTime().getTime() !== o.endTime.getTime();

  try {
    if (needTitle) event.setTitle(o.calendarTitle);
    if (needTime) event.setTime(o.startTime, o.endTime);
    event.setDescription(o.description);
  } catch (e) {
    Logger.log("CalendarApp update not allowed for " + o.dateStr +
      " (" + e.message + ") — falling back to advanced Calendar API.");
    var patch = { summary: o.calendarTitle, description: o.description };
    if (needTime) {
      patch.start = { dateTime: toRfc3339_(o.startTime, o.timeZone), timeZone: o.timeZone };
      patch.end = { dateTime: toRfc3339_(o.endTime, o.timeZone), timeZone: o.timeZone };
    }
    patchCalendarEvent_(calendar, event, patch);
  }
}

/**
 * Formats a Date as an RFC3339 local date-time string (no offset; the timeZone
 * is supplied separately to the Calendar API).
 * @param {Date} date
 * @param {string} timeZone
 * @return {string}
 */
function toRfc3339_(date, timeZone) {
  return Utilities.formatDate(date, timeZone, "yyyy-MM-dd'T'HH:mm:ss");
}

/**
 * Updates a labelled line ("Label: value") in a description string in place.
 * - If value is non-null/non-empty: updates the existing line or appends it.
 * - If value is null/empty: removes the line cleanly.
 * @param {string} desc
 * @param {string} label
 * @param {*} value
 * @return {string} Updated description
 */
function updateDescriptionField_(desc, label, value) {
  var pattern = new RegExp("^" + label + ":.*$", "m");
  var hasValue = value !== null && value !== undefined && String(value).trim() !== "";

  if (hasValue) {
    var line = label + ": " + value;
    if (pattern.test(desc)) {
      return desc.replace(pattern, line);
    } else {
      return desc + "\n" + line;
    }
  } else {
    return desc.replace(pattern, "").replace(/\n{2,}/g, "\n").trim();
  }
}

/**
 * Parses an RSVP count from a cell value like "TTL RSVP =42".
 * @param {*} raw
 * @return {number|null}
 */
function parseRsvpCount_(raw) {
  if (!raw) return null;
  var match = String(raw).match(/=\s*(\d+)/);
  return match ? parseInt(match[1], 10) : null;
}

/**
 * Parses a date value (Date object or string) into a date-only Date.
 * @param {Date|string} value
 * @return {Date|null} Date with time zeroed out, or null if unparseable
 */
function parseEventDate_(value) {
  if (!value) return null;

  if (value instanceof Date) {
    return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }

  if (typeof value === "string") {
    var parsed = new Date(value);
    if (!isNaN(parsed)) {
      return new Date(parsed.getFullYear(), parsed.getMonth(), parsed.getDate());
    }
  }

  return null;
}

/**
 * Builds a Date object at a specific hour/minute in a given timezone.
 * Uses Utilities.parseDate to correctly handle CST/CDT transitions.
 * @param {Date} dateOnly - A date with time zeroed out
 * @param {number} hour - Hour (0-23)
 * @param {number} minute - Minute (0-59)
 * @param {string} timeZone - IANA timezone string (e.g. "America/Chicago")
 * @return {Date}
 */
function buildDateInTimeZone_(dateOnly, hour, minute, timeZone) {
  var year = dateOnly.getFullYear();
  var month = ("0" + (dateOnly.getMonth() + 1)).slice(-2);
  var day = ("0" + dateOnly.getDate()).slice(-2);
  var hourStr = ("0" + hour).slice(-2);
  var minuteStr = ("0" + minute).slice(-2);

  var dateTimeString = year + "-" + month + "-" + day + " " + hourStr + ":" + minuteStr + ":00";
  return Utilities.parseDate(dateTimeString, timeZone, "yyyy-MM-dd HH:mm:ss");
}
