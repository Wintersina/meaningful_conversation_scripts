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
 * TODO: If no match is found, insert a new row with rsvp'd: no / attended: yes
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
  if (!currentValue || currentValue === "-" || currentValue === "--") {
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
  if (String(currentValue).toLowerCase().indexOf("attended: yes") !== -1) {
    return currentValue;
  }

  // Fallback
  return RSVP_DROP_DOWN_CONSTANTS.NO_ATTENDED_YES;
}

/**
 * Preloads Contact List data for matching.
 * Returns an object with arrays indexed by contact row (0-based from ROW_12).
 */
function preloadContactData_(contactListSheet) {
  var lastRow = contactListSheet.getLastRow();
  var lastCol = contactListSheet.getLastColumn();
  var dataStartRow = ROW_NUMBERS.ROW_12;
  var numRows = lastRow - dataStartRow + 1;

  if (numRows <= 0) return null;

  // Read names (col A), first names (col C), last names (col D), emails (col F), phones (col G)
  var names = contactListSheet.getRange(dataStartRow, 1, numRows, 1).getValues();
  var firstNames = contactListSheet.getRange(dataStartRow, 3, numRows, 1).getValues();
  var lastNames = contactListSheet.getRange(dataStartRow, 4, numRows, 1).getValues();
  var emails = contactListSheet.getRange(dataStartRow, 6, numRows, 1).getValues();
  var phones = contactListSheet.getRange(dataStartRow, 7, numRows, 1).getValues();

  // Read event dates from Row 6 and find event columns
  var headerRow5 = contactListSheet.getRange(ROW_NUMBERS.ROW_5, 1, 1, lastCol).getValues()[0];
  var attendedColIndex = headerRow5.indexOf(COL_CONSTANTS.EVENTS_ATTENDED);
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
    attendedColIndex: attendedColIndex,
    lastCol: lastCol
  };
}

/**
 * Finds the event column (1-based) whose Row 6 date matches the signup date.
 * Compares year/month/day only.
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
 *
 * Priority: email → phone → name
 */
function findContactMatch_(contactData, signupName, signupEmail, signupPhone) {
  var normSignupName = normalizeByStrippingWhiteSpaceAtTheEnd(signupName);
  var normSignupEmail = signupEmail ? String(signupEmail).trim().toLowerCase() : "";
  var normSignupPhone = normalizePhone_(signupPhone);

  // Pass 1: Email match (if signup has an email)
  if (normSignupEmail) {
    for (var i = 0; i < contactData.numRows; i++) {
      var contactEmail = String(contactData.emails[i][0] || "").toLowerCase();
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
      if (contactName && contactName === normSignupName) {
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
 * Processes a single signup row against the Contact List.
 * Returns true if a match was found and attendance was marked.
 */
function processSignupRow_(contactListSheet, contactData, signupRow) {
  var signupName = signupRow[SIGNUP_COLS.NAME];
  var signupEmail = signupRow[SIGNUP_COLS.EMAIL];
  var signupPhone = signupRow[SIGNUP_COLS.PHONE];
  var signupTimestamp = signupRow[SIGNUP_COLS.TIMESTAMP];

  if (!signupName) return false;

  // Find the event column by matching signup date to Row 6 dates
  var eventCol = findEventColumnByDate_(contactData, signupTimestamp);
  if (eventCol === -1) {
    Logger.log("No event column found for date: " + signupTimestamp + " (name: " + signupName + ")");
    return false;
  }

  // Find matching contact row
  var contactRow = findContactMatch_(contactData, signupName, signupEmail, signupPhone);
  if (contactRow === -1) {
    // TODO: Insert new row with rsvp'd: no / attended: yes
    Logger.log("NO MATCH FOUND for: " + signupName + " (email: " + signupEmail + ", phone: " + signupPhone + ")");
    return false;
  }

  // Update attendance in the correct event column
  var cell = contactListSheet.getRange(contactRow, eventCol);
  var currentValue = cell.getValue();
  var newValue = markAttendedYes_(currentValue);

  if (newValue !== currentValue) {
    cell.setValue(newValue);
    Logger.log("Marked attended: " + signupName + " → row " + contactRow + ", col " + columnToLetter(eventCol));
  } else {
    Logger.log("Already attended: " + signupName + " → row " + contactRow);
  }

  // Update phone if needed
  updatePhoneIfNeeded_(contactListSheet, contactRow, signupPhone);

  return true;
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
  var unmatched = 0;

  for (var i = 0; i < signupData.length; i++) {
    var row = signupData[i];
    if (!row[SIGNUP_COLS.NAME]) continue; // skip empty rows

    var found = processSignupRow_(contactListSheet, contactData, row);
    if (found) {
      matched++;
    } else {
      unmatched++;
    }
  }

  Logger.log("markAttendanceFromSignupSheet complete. Matched: " + matched + ", Unmatched: " + unmatched);
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

  var [contactListSheet] = sheetsByName();
  var contactData = preloadContactData_(contactListSheet);
  if (!contactData) {
    Logger.log("No contact data to match against");
    return;
  }

  var found = processSignupRow_(contactListSheet, contactData, signupRow);
  if (found) {
    Logger.log("Form submit: match found and attendance marked");
  } else {
    Logger.log("Form submit: no match found for " + signupRow[SIGNUP_COLS.NAME]);
  }
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
