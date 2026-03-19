/**
 * Syncs events from the Contact List sheet to Google Calendar.
 * Idempotent — safe to run multiple times without creating duplicates.
 */
function syncEventsToGoogleCalendar() {
  Logger.log("starting syncEventsToGoogleCalendar");

  var calendar = CalendarApp.getDefaultCalendar();
  var contactListSheet = sheetsByName()[0];

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

  var timeZone = "America/Chicago";
  var createdCount = 0;
  var skippedCount = 0;

  for (var i = 0; i < numCols; i++) {
    var dateValue = dates[i];
    var titleValue = titles[i];

    if (!titleValue || String(titleValue).trim() === "") {
      continue;
    }

    var eventDate = parseEventDate_(dateValue);
    if (!eventDate) {
      Logger.log("Skipping column " + (startCol + i) + ": could not parse date '" + dateValue + "'");
      continue;
    }

    var title = String(titleValue).trim();
    var room = rooms[i] ? String(rooms[i]).trim() : "";
    var rsvpCount = parseRsvpCount_(rsvpRaw[i]);

    // Append room number to title if available
    var calendarTitle = room ? title + " (" + room + ")" : title;

    // Build start (6:30 PM) and end (8:00 PM) in America/Chicago
    var startTime = buildDateInTimeZone_(eventDate, 18, 30, timeZone);
    var endTime = buildDateInTimeZone_(eventDate, 20, 0, timeZone);

    var now = new Date();
    var nowStr = Utilities.formatDate(now, timeZone, "MMM d, yyyy h:mm a");
    var oneMonthAgo = new Date(now.getFullYear(), now.getMonth() - 1, now.getDate());

    if (startTime < oneMonthAgo) {
      Logger.log("Skipping past event (> 1 month old): " + calendarTitle);
      continue;
    }

    // Idempotency check: find existing event by date + title contains
    var existingEvent = findCalendarEvent_(calendar, startTime, endTime, title);
    if (existingEvent) {
      var oldTitle = existingEvent.getTitle();
      var titleChanged = oldTitle !== calendarTitle;

      var desc = existingEvent.getDescription() || "";
      var newDesc = updateDescriptionField_(desc, "Room", room || null);
      newDesc = updateDescriptionField_(newDesc, "RSVPs", rsvpCount !== null ? rsvpCount : null);
      var descChanged = newDesc !== desc;

      if (titleChanged) {
        try {
          Logger.log("Updating title: '" + oldTitle + "' → '" + calendarTitle + "'");
          existingEvent.setTitle(calendarTitle);
        } catch (e) {
          Logger.log("Could not update title for '" + oldTitle + "' — update it manually in Google Calendar: " + e.message);
        }
      }
      if (titleChanged || descChanged) {
        try {
          newDesc = updateDescriptionField_(newDesc, "Last synced", nowStr);
          Logger.log("Updating description for: " + calendarTitle);
          existingEvent.setDescription(newDesc);
        } catch (e) {
          Logger.log("Could not update description for '" + calendarTitle + "': " + e.message);
        }
      }
      skippedCount++;
      continue;
    }

    // Create the event
    var dateStr = Utilities.formatDate(startTime, timeZone, "MMMM d, yyyy");
    var description = "Last synced: " + nowStr;
    if (room) {
      description += "\nRoom: " + room;
    }
    if (rsvpCount !== null) {
      description += "\nRSVPs: " + rsvpCount;
    }
    calendar.createEvent(calendarTitle, startTime, endTime, {
      description: description
    });
    createdCount++;
    Logger.log("Created event: " + calendarTitle + " on " + dateStr);
  }

  Logger.log("syncEventsToGoogleCalendar complete. Created: " + createdCount + ", Skipped (already exist): " + skippedCount);
}

/**
 * Returns the first calendar event in the time window whose title contains the given title,
 * or null if none found.
 * @param {CalendarApp.Calendar} calendar
 * @param {Date} startTime
 * @param {Date} endTime
 * @param {string} title - Base title without room suffix
 * @return {CalendarApp.CalendarEvent|null}
 */
function findCalendarEvent_(calendar, startTime, endTime, title) {
  var events = calendar.getEvents(startTime, endTime);
  var normalizedTitle = normalizeString(title).toLowerCase();

  for (var j = 0; j < events.length; j++) {
    var existingTitle = normalizeString(events[j].getTitle()).toLowerCase();
    if (existingTitle.includes(normalizedTitle)) {
      return events[j];
    }
  }
  return null;
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
 * Mirrors the pattern from insert_new_col_and_row.gs.
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
