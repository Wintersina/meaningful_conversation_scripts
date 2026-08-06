/**
 * Marks attendance in the Contact List based on sign-ins from the
 * "STL MO Barcode (signup sheet)" sheet.
 *
 * Matching priority:
 *   1. Exact email match (substring — handles multi-email cells)
 *   2. Normalized phone match (digits only)
 *   3. Fuzzy name match (normalized, collapsed whitespace, lowercased)
 *
 * When a match is found:
 *   - If the cell has "attended: ?" → flip to "attended: yes" (keep RSVP prefix)
 *   - If the cell is empty / dash → set to "rsvp'd: no\nattended: yes"
 *
 * When NO match is found (walk-in): a new row is inserted into today's event
 * block in the Contact List (two rows below the block title, same placement as
 * the EventBrite import) with rsvp'd: no / attended: yes.
 */

// ── Signup sheet column indices (0-based) ──────────────────────────
var SIGNUP_COLS = {
  TIMESTAMP: 0, // A
  NAME: 1,      // B
  EMAIL: 2,     // C
  PHONE: 3,     // D
  COMMENTS: 4   // E
};

/**
 * Strips all non-digit characters from a phone string for comparison.
 */
function normalizePhone_(phone) {
  if (!phone) return "";
  return String(phone).replace(/\D/g, "");
}

/**
 * Formats a phone number with dashes: 3141234567 → 314-123-4567
 * Returns the original string if it doesn't have exactly 10 digits.
 */
function formatPhoneWithDashes_(phone) {
  var digits = normalizePhone_(phone);
  if (digits.length === 11 && digits.charAt(0) === "1") {
    digits = digits.substring(1); // strip leading 1
  }
  if (digits.length === 10) {
    return digits.substring(0, 3) + "-" + digits.substring(3, 6) + "-" + digits.substring(6);
  }
  return String(phone); // return as-is if not 10 digits
}

/**
 * Given an existing RSVP cell value, returns the "attended: yes" version.
 * Keeps the RSVP prefix (yes/maybe/no) but flips attended to yes.
 * If empty/dash/no value, returns NO_ATTENDED_YES.
 */
function markAttendedYes_(currentValue) {
  // Robust empty detection — handles null, undefined, empty string, whitespace-only, dash
  var trimmed = currentValue == null ? "" : String(currentValue).trim();
  if (!trimmed || trimmed === "-" || trimmed === "--") {
    return RSVP_DROP_DOWN_CONSTANTS.NO_ATTENDED_YES;
  }

  // Map from "attended: ?" variants to "attended: yes" variants
  var attendedMap = {};
  attendedMap[RSVP_DROP_DOWN_CONSTANTS.YES_ATTENDED] = RSVP_DROP_DOWN_CONSTANTS.YES_ATTENDED_YES;
  attendedMap[RSVP_DROP_DOWN_CONSTANTS.MAYBE_ATTENDED] = RSVP_DROP_DOWN_CONSTANTS.MAYBE_ATTENDED_YES;
  attendedMap[RSVP_DROP_DOWN_CONSTANTS.NO_ATTENDED] = RSVP_DROP_DOWN_CONSTANTS.NO_ATTENDED_YES;

  // Also handle "attended: no" → flip to yes
  attendedMap[RSVP_DROP_DOWN_CONSTANTS.YES_ATTENDED_NO] = RSVP_DROP_DOWN_CONSTANTS.YES_ATTENDED_YES;
  attendedMap[RSVP_DROP_DOWN_CONSTANTS.NO_ATTENDED_NO] = RSVP_DROP_DOWN_CONSTANTS.NO_ATTENDED_YES;

  if (attendedMap[currentValue]) {
    return attendedMap[currentValue];
  }

  // Already attended: yes — no change needed
  if (trimmed.toLowerCase().indexOf("attended: yes") !== -1) {
    return currentValue;
  }

  // Fallback: matched person but unrecognized cell value — treat as no RSVP, mark attended
  return RSVP_DROP_DOWN_CONSTANTS.NO_ATTENDED_YES;
}

/**
 * Preloads Contact List data for matching.
 * Uses a single bulk read (cols A-G) instead of 5 separate getRange calls.
 * Returns an object with arrays indexed by contact row (0-based from ROW_12).
 */
function preloadContactData_(contactListSheet) {
  var lastRow = contactListSheet.getLastRow();
  var lastCol = contactListSheet.getLastColumn();
  var dataStartRow = ROW_NUMBERS.ROW_12;
  var numRows = lastRow - dataStartRow + 1;

  if (numRows <= 0) return null;

  // Single bulk read: columns A through G (7 columns) for all contact rows
  var bulkData = contactListSheet.getRange(dataStartRow, 1, numRows, 7).getValues();

  // Extract columns from bulk data (0-based: A=0, C=2, D=3, F=5, G=6)
  var names = [];
  var firstNames = [];
  var lastNames = [];
  var emails = [];
  var phones = [];

  for (var i = 0; i < bulkData.length; i++) {
    names.push([bulkData[i][0]]);
    firstNames.push([bulkData[i][2]]);
    lastNames.push([bulkData[i][3]]);
    emails.push([bulkData[i][5]]);
    phones.push([bulkData[i][6]]);
  }

  // Read event dates from Row 6
  var startCol = HELPER_CONSTANTS.EVENT_NAMES_START_COL; // O = 15
  var eventDates = contactListSheet.getRange(ROW_NUMBERS.ROW_6, startCol, 1, lastCol - startCol + 1).getValues()[0];

  return {
    names: names,
    firstNames: firstNames,
    lastNames: lastNames,
    emails: emails,
    phones: phones,
    eventDates: eventDates,
    startCol: startCol,
    dataStartRow: dataStartRow,
    numRows: numRows,
    lastCol: lastCol
  };
}

/**
 * Lightweight version for the form trigger — finds today's event column,
 * then only loads contact data for rows that have a value in that column.
 * This avoids scanning all 3000+ rows for every form submission.
 *
 * Returns null if no event column matches the date, or an object with:
 *   - candidates: array of { sheetRow, name, firstName, lastName, email, phone, rsvpValue }
 *   - resolvedEventCol: the 1-based column index for today's event
 */
function preloadCandidatesForDate_(contactListSheet, targetDate) {
  var lastRow = contactListSheet.getLastRow();
  var lastCol = contactListSheet.getLastColumn();
  var dataStartRow = ROW_NUMBERS.ROW_12;
  var numRows = lastRow - dataStartRow + 1;

  if (numRows <= 0) return null;

  // Find the event column for this specific date
  var startCol = HELPER_CONSTANTS.EVENT_NAMES_START_COL;
  var eventDates = contactListSheet.getRange(ROW_NUMBERS.ROW_6, startCol, 1, lastCol - startCol + 1).getValues()[0];

  var eventCol = -1;
  if (targetDate instanceof Date && !isNaN(targetDate.getTime())) {
    var tY = targetDate.getFullYear();
    var tM = targetDate.getMonth();
    var tD = targetDate.getDate();

    for (var i = 0; i < eventDates.length; i++) {
      var d = eventDates[i];
      if (d instanceof Date && d.getFullYear() === tY && d.getMonth() === tM && d.getDate() === tD) {
        eventCol = startCol + i;
        break;
      }
    }
  }

  if (eventCol === -1) return null;

  // Read today's event column to find which rows have RSVP values
  var eventColValues = contactListSheet.getRange(dataStartRow, eventCol, numRows, 1).getValues();

  // Collect row indices that have any value in the event column
  var candidateIndices = []; // 0-based offsets from dataStartRow
  for (var i = 0; i < eventColValues.length; i++) {
    var val = eventColValues[i][0];
    if (val && val !== "-" && val !== "--") {
      candidateIndices.push(i);
    }
  }

  // Also read all contact data (A-G) in one call, but only build candidate objects
  // for rows that have an RSVP value — keeps the search list small
  var bulkData = contactListSheet.getRange(dataStartRow, 1, numRows, 7).getValues();

  var candidates = [];
  for (var j = 0; j < candidateIndices.length; j++) {
    var idx = candidateIndices[j];
    candidates.push({
      sheetRow: dataStartRow + idx, // 1-based
      name: bulkData[idx][0],
      firstName: bulkData[idx][2],
      lastName: bulkData[idx][3],
      email: bulkData[idx][5],
      phone: bulkData[idx][6],
      rsvpValue: eventColValues[idx][0]
    });
  }

  // Also keep the full bulk data for fallback matching (someone who showed up
  // but has no RSVP — their cell is empty, so they won't be in candidates)
  var allRows = [];
  for (var i = 0; i < bulkData.length; i++) {
    allRows.push({
      sheetRow: dataStartRow + i,
      name: bulkData[i][0],
      firstName: bulkData[i][2],
      lastName: bulkData[i][3],
      email: bulkData[i][5],
      phone: bulkData[i][6]
    });
  }

  return {
    candidates: candidates,
    allRows: allRows,
    resolvedEventCol: eventCol
  };
}

/**
 * Finds the event column (1-based) whose Row 6 date matches the signup date.
 * Compares year/month/day only. Used by the batch/manual flow.
 */
function findEventColumnByDate_(contactData, signupDate) {
  if (!(signupDate instanceof Date)) {
    signupDate = new Date(signupDate);
  }
  if (isNaN(signupDate.getTime())) return -1;

  var targetY = signupDate.getFullYear();
  var targetM = signupDate.getMonth();
  var targetD = signupDate.getDate();

  for (var i = 0; i < contactData.eventDates.length; i++) {
    var d = contactData.eventDates[i];
    if (!(d instanceof Date)) continue;

    if (d.getFullYear() === targetY && d.getMonth() === targetM && d.getDate() === targetD) {
      return contactData.startCol + i; // 1-based column index
    }
  }

  return -1;
}

/**
 * Finds the best matching contact row for a signup entry.
 * Returns the 1-based sheet row number, or -1 if no match.
 * Used by the batch/manual flow (scans all rows).
 *
 * Priority: email → phone → name
 */
function findContactMatch_(contactData, signupName, signupEmail, signupPhone) {
  var normSignupName = normalizeByStrippingWhiteSpaceAtTheEnd(signupName);
  var normSignupEmail = signupEmail ? String(signupEmail).trim().toLowerCase() : "";
  var normSignupPhone = normalizePhone_(signupPhone);
  // Spaceless version for "MaryBeth" vs "mary beth" matching
  var spacelessSignupName = normSignupName ? normSignupName.replace(/\s/g, "") : "";

  // Pass 1: Email match (if signup has an email)
  if (normSignupEmail) {
    for (var i = 0; i < contactData.numRows; i++) {
      var contactEmail = String(contactData.emails[i][0] == null ? "" : contactData.emails[i][0]).trim().toLowerCase();
      if (contactEmail && contactEmail.indexOf(normSignupEmail) !== -1) {
        return contactData.dataStartRow + i;
      }
    }
  }

  // Pass 2: Phone match (if signup has a phone)
  if (normSignupPhone) {
    for (var i = 0; i < contactData.numRows; i++) {
      var contactPhone = normalizePhone_(contactData.phones[i][0]);
      if (contactPhone && contactPhone === normSignupPhone) {
        return contactData.dataStartRow + i;
      }
    }
  }

  // Pass 3: Fuzzy name match
  if (normSignupName) {
    for (var i = 0; i < contactData.numRows; i++) {
      var contactName = normalizeByStrippingWhiteSpaceAtTheEnd(contactData.names[i][0]);
      if (!contactName) continue;

      // Exact normalized match
      if (contactName === normSignupName) {
        return contactData.dataStartRow + i;
      }

      // Spaceless match: "marybeth" === "marybeth" even if one was "mary beth"
      if (spacelessSignupName && contactName.replace(/\s/g, "") === spacelessSignupName) {
        return contactData.dataStartRow + i;
      }
    }

    // Pass 3b: Try first+last separately (handles name order differences)
    var signupParts = normSignupName.split(" ");
    if (signupParts.length >= 2) {
      var signupFirst = signupParts[0];
      var signupLast = signupParts[signupParts.length - 1];

      for (var i = 0; i < contactData.numRows; i++) {
        var contactFirst = normalizeByStrippingWhiteSpaceAtTheEnd(contactData.firstNames[i][0]);
        var contactLast = normalizeByStrippingWhiteSpaceAtTheEnd(contactData.lastNames[i][0]);

        if (!contactFirst || !contactLast) continue;

        // Check both orderings
        if ((contactFirst === signupFirst && contactLast === signupLast) ||
            (contactFirst === signupLast && contactLast === signupFirst)) {
          return contactData.dataStartRow + i;
        }
      }
    }
  }

  return -1;
}

/**
 * Searches a list of row objects for a match against signup data.
 * Used by the optimized trigger flow.
 * Returns the matching row object, or null.
 *
 * Priority: email → phone → name
 */
function findMatchInList_(rows, signupName, signupEmail, signupPhone) {
  var normSignupName = normalizeByStrippingWhiteSpaceAtTheEnd(signupName);
  var normSignupEmail = signupEmail ? String(signupEmail).trim().toLowerCase() : "";
  var normSignupPhone = normalizePhone_(signupPhone);
  var spacelessSignupName = normSignupName ? normSignupName.replace(/\s/g, "") : "";

  // Pass 1: Email
  if (normSignupEmail) {
    for (var i = 0; i < rows.length; i++) {
      var contactEmail = String(rows[i].email == null ? "" : rows[i].email).trim().toLowerCase();
      if (contactEmail && contactEmail.indexOf(normSignupEmail) !== -1) {
        return rows[i];
      }
    }
  }

  // Pass 2: Phone
  if (normSignupPhone) {
    for (var i = 0; i < rows.length; i++) {
      var contactPhone = normalizePhone_(rows[i].phone);
      if (contactPhone && contactPhone === normSignupPhone) {
        return rows[i];
      }
    }
  }

  // Pass 3: Name
  if (normSignupName) {
    for (var i = 0; i < rows.length; i++) {
      var contactName = normalizeByStrippingWhiteSpaceAtTheEnd(rows[i].name);
      if (!contactName) continue;

      // Exact normalized match
      if (contactName === normSignupName) {
        return rows[i];
      }

      // Spaceless match: "marybeth" === "marybeth"
      if (spacelessSignupName && contactName.replace(/\s/g, "") === spacelessSignupName) {
        return rows[i];
      }
    }

    // Pass 3b: First+last swap
    var parts = normSignupName.split(" ");
    if (parts.length >= 2) {
      var sFirst = parts[0];
      var sLast = parts[parts.length - 1];

      for (var i = 0; i < rows.length; i++) {
        var cFirst = normalizeByStrippingWhiteSpaceAtTheEnd(rows[i].firstName);
        var cLast = normalizeByStrippingWhiteSpaceAtTheEnd(rows[i].lastName);
        if (!cFirst || !cLast) continue;

        if ((cFirst === sFirst && cLast === sLast) ||
            (cFirst === sLast && cLast === sFirst)) {
          return rows[i];
        }
      }
    }
  }

  return null;
}

/**
 * Updates the phone number in the Contact List if the signup has a different one.
 * Appends with comma if existing phone differs. Normalizes to dash format.
 */
function updatePhoneIfNeeded_(contactListSheet, contactRow, signupPhone) {
  if (!signupPhone) return;

  var normSignupPhone = normalizePhone_(signupPhone);
  if (!normSignupPhone) return;

  var existingPhone = String(contactListSheet.getRange(contactRow, 7).getValue() || "");
  var normExistingPhone = normalizePhone_(existingPhone);

  // No existing phone — just set it
  if (!normExistingPhone) {
    contactListSheet.getRange(contactRow, 7).setValue(formatPhoneWithDashes_(signupPhone));
    return;
  }

  // Phone already contains this number
  if (normExistingPhone.indexOf(normSignupPhone) !== -1) return;

  // Different phone — append
  var formatted = formatPhoneWithDashes_(signupPhone);
  contactListSheet.getRange(contactRow, 7).setValue(existingPhone + ", " + formatted);
}

/**
 * Splits a full name into [first, last]. First token → first name,
 * everything after → last name (may be empty for single-word names).
 */
function splitSignupName_(fullName) {
  var trimmed = String(fullName || "").trim().replace(/\s+/g, " ");
  if (!trimmed) return ["", ""];
  var spaceIdx = trimmed.indexOf(" ");
  if (spaceIdx === -1) return [trimmed, ""];
  return [trimmed.substring(0, spaceIdx), trimmed.substring(spaceIdx + 1)];
}

/**
 * Inserts a walk-in row into the Contact List for someone who signed in via
 * barcode but has no RSVP anywhere in the sheet.
 *
 * Placement mirrors moveRowsFromEventBriteImportToContactList: the row goes
 * TWO rows below the event-block title in Column B that matches today's event
 * title (Row 7 of the resolved event column).
 *
 * Returns the inserted 1-based row number, or -1 if the event block could not
 * be located (nothing is inserted in that case).
 */
function insertWalkInContactRow_(contactListSheet, eventCol, signupName, signupEmail, signupPhone, signupTimestamp) {
  var eventTitle = String(contactListSheet.getRange(ROW_NUMBERS.ROW_7, eventCol).getValue() || "").trim();
  if (!eventTitle) {
    Logger.log("Walk-in insert skipped: no event title in row 7, col " + columnToLetter(eventCol));
    return -1;
  }

  // Find the matching event block title in Column B (bottom-most wins, like the EB import)
  var lastRow = contactListSheet.getLastRow();
  var colB = contactListSheet.getRange(1, 2, lastRow, 1).getValues();
  var matchRow = -1;
  var normTitle = normalizeString(eventTitle);
  for (var i = colB.length - 1; i >= 0; i--) {
    if (normalizeString(String(colB[i][0] || "")).trim() === normTitle) {
      matchRow = i + 1; // 1-based
      break;
    }
  }
  if (matchRow === -1) {
    Logger.log("Walk-in insert skipped: event block '" + eventTitle + "' not found in Column B");
    return -1;
  }

  var rsvpCol = findColMarker_(contactListSheet, MARKER_KEYS.EVENTS_RSVPD, COL_CONSTANTS.EVENTS_RSVPD);
  var attendedCol = findColMarker_(contactListSheet, MARKER_KEYS.EVENTS_ATTENDED, COL_CONSTANTS.EVENTS_ATTENDED);
  if (rsvpCol === -1 || attendedCol === -1) {
    Logger.log("Walk-in insert skipped: counter columns not found");
    return -1;
  }
  var attendedColLetter = columnToLetter(attendedCol);

  var insertAt = matchRow + 2; // two rows below the block title

  // Make sure we don't insert inside a collapsed/hidden group
  expandRowGroupsAtRow_(contactListSheet, insertAt);
  contactListSheet.showRows(insertAt);

  contactListSheet.insertRowsBefore(insertAt, 1);

  // Identity columns: C=first, D=last, F=email, G=phone
  var nameParts = splitSignupName_(signupName);
  contactListSheet.getRange(insertAt, 3, 1, 2).setValues([[nameParts[0], nameParts[1]]]);
  if (signupEmail) {
    contactListSheet.getRange(insertAt, COLUMN_INDEX.EMAIL + 1).setValue(String(signupEmail).trim());
  }
  if (signupPhone) {
    contactListSheet.getRange(insertAt, COLUMN_INDEX.PHONE + 1).setValue(formatPhoneWithDashes_(signupPhone));
  }

  // Origin columns: J=signup date/time, K=platform, L=event title
  if (signupTimestamp) {
    contactListSheet.getRange(insertAt, COLUMN_INDEX.SIGNUP_DATE_TIME + 1).setValue(signupTimestamp);
  }
  contactListSheet.getRange(insertAt, COLUMN_INDEX.SIGNUP_PLATFORM + 1).setValue("Walk-in");
  contactListSheet.getRange(insertAt, COLUMN_INDEX.SIGNUP_EVENT_TITLE + 1).setValue(eventTitle);

  // Formulas: A=name concat, counter columns
  contactListSheet.getRange(insertAt, 1).setFormula(FORMULAS.CONCATENATE_NAME(insertAt));
  contactListSheet.getRange(insertAt, attendedCol).setFormula(FORMULAS.COUNT_ATTENDED(insertAt, attendedColLetter));
  contactListSheet.getRange(insertAt, rsvpCol).setFormula(FORMULAS.COUNT_RSVP(insertAt, attendedColLetter));

  // Today's event cell: dropdown + rsvp'd: no / attended: yes
  var rsvpValidationRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(Object.values(RSVP_DROP_DOWN_CONSTANTS), true)
    .setAllowInvalid(false)
    .build();
  var eventCell = contactListSheet.getRange(insertAt, eventCol);
  eventCell.setDataValidation(rsvpValidationRule);
  eventCell.setValue(RSVP_DROP_DOWN_CONSTANTS.NO_ATTENDED_YES);

  // Dash-fill the remaining event columns to the right (same as EB import)
  var dashStart = eventCol + 1;
  var dashEnd = attendedCol - 1;
  if (dashStart <= dashEnd) {
    var width = dashEnd - dashStart + 1;
    contactListSheet.getRange(insertAt, dashStart, 1, width).setValues([Array(width).fill("-")]);
  }

  // Static formatting
  contactListSheet
    .getRange(insertAt, 1, 1, contactListSheet.getLastColumn())
    .setFontSize(UI_CONSTANTS.FONT_SIZE)
    .setFontFamily(UI_CONSTANTS.FONT_STYLE)
    .setHorizontalAlignment(UI_CONSTANTS.ALIGNMENT_CENTER);

  Logger.log("Inserted walk-in row: " + signupName + " → row " + insertAt + " (block '" + eventTitle + "')");
  return insertAt;
}

/**
 * Processes a single signup row against the Contact List.
 * Returns "matched", "inserted", or "skipped".
 */
function processSignupRow_(contactListSheet, contactData, signupRow) {
  var signupName = signupRow[SIGNUP_COLS.NAME];
  var signupEmail = signupRow[SIGNUP_COLS.EMAIL];
  var signupPhone = signupRow[SIGNUP_COLS.PHONE];
  var signupTimestamp = signupRow[SIGNUP_COLS.TIMESTAMP];

  if (!signupName) return "skipped";

  // Find the event column by matching signup date to Row 6 dates
  var eventCol = findEventColumnByDate_(contactData, signupTimestamp);
  if (eventCol === -1) {
    Logger.log("No event column found for date: " + signupTimestamp + " (name: " + signupName + ")");
    return "skipped";
  }

  // Find matching contact row
  var contactRow = findContactMatch_(contactData, signupName, signupEmail, signupPhone);
  if (contactRow === -1) {
    // Walk-in: barcode sign-in with no RSVP anywhere — insert a fresh row
    Logger.log("NO MATCH FOUND for: " + signupName + " (email: " + signupEmail + ", phone: " + signupPhone + ") — inserting walk-in row");
    var insertedRow = insertWalkInContactRow_(contactListSheet, eventCol, signupName, signupEmail, signupPhone, signupTimestamp);
    return insertedRow === -1 ? "skipped" : "inserted";
  }

  // Update attendance in the correct event column
  var cell = contactListSheet.getRange(contactRow, eventCol);
  var currentValue = cell.getValue();
  var currentTrimmed = currentValue == null ? "" : String(currentValue).trim();
  var newValue = markAttendedYes_(currentValue);

  Logger.log("Cell value for " + signupName + " at row " + contactRow + ", col " + columnToLetter(eventCol) + ": [" + currentTrimmed + "] → [" + newValue + "]");

  if (currentTrimmed.toLowerCase().indexOf("attended: yes") !== -1) {
    Logger.log("Already attended: " + signupName + " → row " + contactRow);
  } else {
    cell.setValue(newValue);
    Logger.log("Marked attended: " + signupName + " → row " + contactRow + ", col " + columnToLetter(eventCol));
  }

  // Update phone if needed
  updatePhoneIfNeeded_(contactListSheet, contactRow, signupPhone);

  return "matched";
}

// ═══════════════════════════════════════════════════════════════════
// Public entry points
// ═══════════════════════════════════════════════════════════════════

/**
 * Manual button: batch-processes ALL rows in the signup sheet.
 * Useful for backfill or catching up.
 */
function markAttendanceFromSignupSheet() {
  Logger.log("starting markAttendanceFromSignupSheet");

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var signupSheet = ss.getSheetByName(SHEET_NAMES.SIGNUP);
  if (!signupSheet) {
    Logger.log("Signup sheet not found: " + SHEET_NAMES.SIGNUP);
    return;
  }

  var [contactListSheet] = sheetsByName();
  var contactData = preloadContactData_(contactListSheet);
  if (!contactData) {
    Logger.log("No contact data to match against");
    return;
  }

  // Read all signup rows (skip header row 1)
  var lastRow = signupSheet.getLastRow();
  if (lastRow < 2) {
    Logger.log("No signup data found");
    return;
  }

  var lastCol = signupSheet.getLastColumn();
  var signupData = signupSheet.getRange(2, 1, lastRow - 1, Math.min(lastCol, 5)).getValues();

  var matched = 0;
  var inserted = 0;
  var skipped = 0;

  for (var i = 0; i < signupData.length; i++) {
    var row = signupData[i];
    if (!row[SIGNUP_COLS.NAME]) continue; // skip empty rows

    var result = processSignupRow_(contactListSheet, contactData, row);
    if (result === "matched") {
      matched++;
    } else if (result === "inserted") {
      inserted++;
      // Inserting a row shifts everything below it — reload the cached
      // contact data so later signups match against fresh row numbers
      // (and so a duplicate signup for the same walk-in now matches).
      contactData = preloadContactData_(contactListSheet);
    } else {
      skipped++;
    }
  }

  Logger.log("markAttendanceFromSignupSheet complete. Matched: " + matched + ", Walk-ins inserted: " + inserted + ", Skipped: " + skipped);
}

/**
 * Form submit trigger: processes the newly submitted row in real-time.
 * Set up with:
 *   ScriptApp.newTrigger('onSignupFormSubmit')
 *     .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
 *     .onFormSubmit()
 *     .create();
 */
function onSignupFormSubmit(e) {
  Logger.log("onSignupFormSubmit triggered");
  Logger.log("Event object keys: " + (e ? Object.keys(e).join(", ") : "null"));

  if (!e) {
    Logger.log("No event object received");
    return;
  }

  // Use e.range (more reliable) to read the actual row from the signup sheet
  // e.values can be unreliable with some form configurations
  var signupRow;

  if (e.range) {
    var sheet = e.range.getSheet();
    var row = e.range.getRow();
    Logger.log("Form submitted to sheet: " + sheet.getName() + ", row: " + row);

    // Only process if the submission landed in our signup sheet
    if (sheet.getName() !== SHEET_NAMES.SIGNUP) {
      Logger.log("Form submit was to a different sheet (" + sheet.getName() + "), skipping");
      return;
    }

    // Read the full row (columns A-E)
    signupRow = sheet.getRange(row, 1, 1, 5).getValues()[0];
  } else if (e.values) {
    Logger.log("Falling back to e.values");
    signupRow = e.values;
  } else {
    Logger.log("No usable event data (no e.range or e.values)");
    return;
  }

  Logger.log("Signup row: name=" + signupRow[SIGNUP_COLS.NAME] +
    ", email=" + signupRow[SIGNUP_COLS.EMAIL] +
    ", phone=" + signupRow[SIGNUP_COLS.PHONE] +
    ", timestamp=" + signupRow[SIGNUP_COLS.TIMESTAMP]);

  // Parse the signup date to find today's event column directly
  var signupTimestamp = signupRow[SIGNUP_COLS.TIMESTAMP];
  var signupDate = (signupTimestamp instanceof Date) ? signupTimestamp : new Date(signupTimestamp);

  if (isNaN(signupDate.getTime())) {
    Logger.log("Could not parse signup timestamp: " + signupTimestamp);
    return;
  }

  var signupName = signupRow[SIGNUP_COLS.NAME];
  var signupEmail = signupRow[SIGNUP_COLS.EMAIL];
  var signupPhone = signupRow[SIGNUP_COLS.PHONE];

  if (!signupName) {
    Logger.log("No name in signup row — skipping");
    return;
  }

  var [contactListSheet] = sheetsByName();
  var data = preloadCandidatesForDate_(contactListSheet, signupDate);
  if (!data) {
    Logger.log("No event column found for date: " + signupDate + " — skipping");
    return;
  }

  var eventCol = data.resolvedEventCol;

  // First: try to match against only rows that have an RSVP for today's event
  // (small list — typically 10-30 people)
  var match = findMatchInList_(data.candidates, signupName, signupEmail, signupPhone);

  if (match) {
    var newValue = markAttendedYes_(match.rsvpValue);
    if (newValue !== match.rsvpValue) {
      contactListSheet.getRange(match.sheetRow, eventCol).setValue(newValue);
      Logger.log("Marked attended (from candidates): " + signupName + " → row " + match.sheetRow);
    } else {
      Logger.log("Already attended: " + signupName + " → row " + match.sheetRow);
    }
    updatePhoneIfNeeded_(contactListSheet, match.sheetRow, signupPhone);
    return;
  }

  // Fallback: search ALL rows (person exists in contact list but has no RSVP for today)
  var fallbackMatch = findMatchInList_(data.allRows, signupName, signupEmail, signupPhone);

  if (fallbackMatch) {
    var currentValue = contactListSheet.getRange(fallbackMatch.sheetRow, eventCol).getValue();
    var newValue = markAttendedYes_(currentValue);
    contactListSheet.getRange(fallbackMatch.sheetRow, eventCol).setValue(newValue);
    Logger.log("Marked attended (no prior RSVP): " + signupName + " → row " + fallbackMatch.sheetRow);
    updatePhoneIfNeeded_(contactListSheet, fallbackMatch.sheetRow, signupPhone);
    return;
  }

  // Walk-in: barcode sign-in with no matching contact anywhere — insert a fresh row
  Logger.log("NO MATCH FOUND for: " + signupName + " (email: " + signupEmail + ", phone: " + signupPhone + ") — inserting walk-in row");
  insertWalkInContactRow_(contactListSheet, eventCol, signupName, signupEmail, signupPhone, signupTimestamp);
}

/**
 * One-time setup: creates the onFormSubmit trigger.
 * Run this once from the script editor.
 */
function setupSignupFormTrigger() {
  // Remove any existing triggers for this function to avoid duplicates
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "onSignupFormSubmit") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  ScriptApp.newTrigger("onSignupFormSubmit")
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onFormSubmit()
    .create();

  Logger.log("Signup form trigger created successfully");
}
