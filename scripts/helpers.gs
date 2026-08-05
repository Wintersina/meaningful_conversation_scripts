const SHEET_NAMES = {
  SCHEDULE: "Schedule",
  EVENTBRITE: "EventBrite Import",
  CONTACT_LIST: "Contact List"
};
SHEET_NAMES.SIGNUP = "STL MO Barcode (signup sheet)";

const ERROR_MESSAGES = {
  CONTACT_LIST_NOT_FOUND: "'Contact List' sheet not found.",
  SCHEDULE_NOT_FOUND: "'Schedule' sheet not found.",
  EVENTBRITE_NOT_FOUND: "'EventBrite Import' sheet not found."
};

const UI_CONSTANTS = {
  GRAY_BACKGROUND: "#D9D9D9",
  BORDER_COLOR: "black",
  BORDER_STYLE: SpreadsheetApp.BorderStyle.SOLID_THICK,
  WRAP_STRATEGY: SpreadsheetApp.WrapStrategy.OVERFLOW,
  ALIGNMENT_CENTER: "center",
  NUMBER_FORMAT: "0",
  FONT_STYLE:"EB Garamond",
  FONT_SIZE:11,
};


const COL_CONSTANTS = {
  EVENTS_ATTENDED : "# Events Attended",
  EVENTS_RSVPD:"# Events RSVP'd",
  NEXT_STEPS:"Next Steps",
  TOTAL_RSVPD: "Total RSVP'd",
  TOTAL_ATTENDED: "Total Attended",
  TOTAL_ATTENDED_NO_RSVP: "Total Attended w/o RSVP",
  STOP_RSVP: "Stop RSVP",
  RSVP_2_PLUS: "RSVP 2+",
  ATTENDED_PLUS_ONE:"ATTENDED 1+",
  STOP_EMAIL:"Stop Email",
  EMAIL_START:"Start Email"
}

// 0-based column indices into the Contact List row arrays (data[i][n]).
// Mirrors the sheet layout A..N so merge/skip logic reads by name, not magic numbers.
const COLUMN_INDEX = {
  FULL_NAME_KEY: 0,        // A — =CONCATENATE(First, " ", Last); the merge key
  NAME_DISPLAY: 1,         // B — full-name display / total-row label
  FIRST_NAME: 2,           // C
  LAST_NAME: 3,            // D
  STATUS: 4,               // E
  EMAIL: 5,                // F
  PHONE: 6,                // G
  NEIGHBORHOOD: 7,         // H
  AGE: 8,                  // I — AGE? (approximate)
  SIGNUP_DATE_TIME: 9,     // J — Original Signup Date & Time
  SIGNUP_PLATFORM: 10,     // K — Original Signup Platform (FB/MU/EB/etc.)
  SIGNUP_EVENT_TITLE: 11,  // L — Original Signup Event Title
  SIGNUP_EVENT_CODE: 12,   // M — Original Signup Event Code
  DESCRIPTION: 13          // N
};

const RSVP_DROP_DOWN_CONSTANTS = {
  DASH: "-",
  DOUBLE_DASH: "--",
  YES_ATTENDED: "rsvp'd: yes\nattended: ?",
  MAYBE_ATTENDED: "rsvp'd: maybe\nattended: ?",
  NO_ATTENDED: "rsvp'd: no\nattended: ?",
  YES_ATTENDED_YES:"rsvp'd: yes\nattended: yes",
  NO_ATTENDED_YES:"rsvp'd: no\nattended: yes",
  MAYBE_ATTENDED_YES:"rsvp'd: maybe\nattended: yes",
  DASH_ATTENDED_YES:"rsvp'd: -\nattended: yes",
  YES_ATTENDED_NO:"rsvp'd: yes\nattended: no",
  NO_ATTENDED_NO:"rsvp'd: no\nattended: no"
}

const HELPER_CONSTANTS = {
  EVENTBRITE_COLUMN_COUNT: 12,
  CONTACT_LIST_OFFSET: 3,
  FIRST_DATA_ROW: 2,
  TITLE_ROW: 7,
  MAX_ROW_CHECK: 3000,
  EVENT_NAMES_START_COL: 15
};

const FORMULAS = {
  CONCATENATE_NAME: (row) => `=IF(ISTEXT(C${row}),CONCATENATE(C${row}, " ", D${row}),"")`,
  COUNT_ATTENDED: (row, colLetter) => `=IFERROR(COUNTIF(O${row}:${colLetter}${row},"*Attended: yes"),"")`,
  COUNT_RSVP: (row, colLetter) => `=IFERROR(COUNTIF(O${row}:${colLetter}${row},"RSVP'd: yes*"),"")`,
  V_LOOKUP:(colLetter) => `=VLOOKUP(${colLetter}${ROW_NUMBERS.ROW_7},'Event IDs'!$B:$C,2,FALSE)`,
  TTL_RSVP:(colLetter) => `="TTL RSVP =" & COUNTIF(${colLetter}12:${colLetter}3000, "*RSVP'D: yes*")`,
  TTL_ATTND:(colLetter) => `="TTL ATTND =" & COUNTIF(${colLetter}12:${colLetter}3000, "*ATTENDED: yes*")`,
  // Summary formulas for the Total rows at the bottom of each event block
  TOTAL_RSVPD_FORMULA:(colLetter, startRow, endRow) => `=IFERROR(COUNTIF(${colLetter}${startRow}:${colLetter}${endRow},"RSVP'd: yes*"),"")`,
  TOTAL_ATTENDED_FORMULA:(colLetter, startRow, endRow) => `=IFERROR(COUNTIF(${colLetter}${startRow}:${colLetter}${endRow},"*Attended: yes"),"")`,
  TOTAL_ATTENDED_NO_RSVP_FORMULA:(colLetter, startRow, endRow) => `=IFERROR(COUNTIF(${colLetter}${startRow}:${colLetter}${endRow},"*rsvp'd: no\nattended: yes"),"")`,
};


const ROW_NUMBERS = {
  ROW_1 : 1,
  ROW_2 : 2,
  ROW_3 : 3,
  ROW_4 : 4,
  ROW_5 : 5, // very import row
  ROW_6 : 6,
  ROW_7 : 7,
  ROW_8 : 8,
  ROW_9 : 9,
  ROW_10:10,
  ROW_11 : 11,
  ROW_12 : 12,

}

const EMAILER_KEYS = {
  docId : "1r-FDCCevUEO6U7sBXDRy-Mgu-2skMxTilv3cn9AOe8E",
  attachmentId : "1sNvF_oeejJVG3PlnKo5YKepaBrwLfB3A"
}


const SCHEDULE_SHEET_CONSTANTS = {
  SKIP_WORDS : ["OFF", "TBD", "TOPIC"]
}

function columnToLetter(columnIndex) {
  let columnLetter = '';
  while (columnIndex > 0) {
    let remainder = (columnIndex - 1) % 26;
    columnLetter = String.fromCharCode(65 + remainder) + columnLetter;
    columnIndex = Math.floor((columnIndex - 1) / 26);
  }
  return columnLetter;
}

function normalizeString(str) {
  return str.replace(/[‘’‚‛′‵]/g, "'"); // Normalize apostrophes
}

function normalizeByStrippingWhiteSpaceAtTheEnd(str) {
  if (!str) return null;
  // Convert to string, trim whitespace from ends, collapse internal spaces, and lowercase
  str = normalizeString(str)
  return String(str).trim().replace(/\s+/g, " ").toLowerCase();
}

function sheetsByName() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const scheduleSheet = ss.getSheetByName(SHEET_NAMES.SCHEDULE);
  const eventbriteSheet = ss.getSheetByName(SHEET_NAMES.EVENTBRITE);
  const contactListSheet = ss.getSheetByName(SHEET_NAMES.CONTACT_LIST);

  if (!contactListSheet) {
    throw new Error(ERROR_MESSAGES.CONTACT_LIST_NOT_FOUND);
  }
  if (!eventbriteSheet) {
    throw new Error(ERROR_MESSAGES.EVENTBRITE_NOT_FOUND);
  }
  if (!scheduleSheet) {
    throw new Error(ERROR_MESSAGES.SCHEDULE_NOT_FOUND);
  }

  return [contactListSheet, eventbriteSheet, scheduleSheet];
}

// ============================================================================
// DeveloperMetadata marker helpers
// ----------------------------------------------------------------------------
// Landmarks — counter columns in row 5 (# Events Attended/RSVP'd, Next Steps)
// and section-boundary rows (Total RSVP'd, ...) — are located via
// DeveloperMetadata for an O(1) lookup that also auto-shifts when rows/columns
// are inserted. Every lookup is self-healing: if the pin is missing or no
// longer sits on the expected label, it falls back to the classic scan and
// re-pins. This means a sheet with no metadata yet behaves exactly like before
// on the first run, then gets fast afterwards.
// ============================================================================

const MARKER_KEYS = {
  EVENTS_ATTENDED: "mc.col.eventsAttended", // row-5 header
  EVENTS_RSVPD:    "mc.col.eventsRsvpd",     // row-5 header
  NEXT_STEPS:      "mc.col.nextSteps",       // row-5 header
  TOTAL_RSVPD:     "mc.row.totalRsvpd",      // pointer to the CURRENT (last) block
};

const DM_VISIBILITY = SpreadsheetApp.DeveloperMetadataVisibility.DOCUMENT;

/**
 * Find a counter column (row-5 header) by metadata, self-healing to an indexOf
 * scan of row 5. Returns a 1-based column index, or -1 if the label is absent.
 */
function findColMarker_(sheet, key, label) {
  const found = sheet.createDeveloperMetadataFinder()
    .withLocationType(SpreadsheetApp.DeveloperMetadataLocationType.COLUMN)
    .withKey(key).find();

  if (found.length) {
    const col = found[0].getLocation().getColumn().getColumn();
    if (String(sheet.getRange(ROW_NUMBERS.ROW_5, col).getValue()).trim() === label) {
      return col; // pin still valid
    }
    found.forEach(m => m.remove()); // stale pin — drop and rescan
  }

  const lastCol = sheet.getLastColumn();
  const header = sheet.getRange(ROW_NUMBERS.ROW_5, 1, 1, lastCol).getValues()[0];
  const idx0 = header.findIndex(v => String(v).trim() === label);
  if (idx0 === -1) return -1;

  const col = idx0 + 1;
  pinColMarker_(sheet, key, col);
  return col;
}

function pinColMarker_(sheet, key, colIndex) {
  sheet.createDeveloperMetadataFinder()
    .withLocationType(SpreadsheetApp.DeveloperMetadataLocationType.COLUMN)
    .withKey(key).find().forEach(m => m.remove());
  // Metadata requires an *entire* column — a bounded getRange(...maxRows...) is
  // rejected as an "arbitrary range", so use A1 whole-column notation ("P:P").
  const letter = columnToLetter(colIndex);
  sheet.getRange(letter + ":" + letter).addDeveloperMetadata(key, DM_VISIBILITY);
}

/**
 * Find a section-boundary row by metadata, self-healing to a scan of `scanCol`
 * (1-based). Returns a 1-based row index, or -1 if the label is absent. When
 * findLast is true the fallback scan matches the LAST occurrence (used for the
 * repeated "Total RSVP'd" label, where we want the current/last block).
 */
function findRowMarker_(sheet, key, label, scanCol, findLast) {
  const found = sheet.createDeveloperMetadataFinder()
    .withLocationType(SpreadsheetApp.DeveloperMetadataLocationType.ROW)
    .withKey(key).find();

  if (found.length) {
    const row = found[0].getLocation().getRow().getRow();
    if (String(sheet.getRange(row, scanCol).getValue()).trim() === label) {
      return row; // pin still valid
    }
    found.forEach(m => m.remove()); // stale pin — drop and rescan
  }

  const row = scanColumnForLabel_(sheet, label, scanCol, findLast);
  if (row > 0) pinRowMarker_(sheet, key, row);
  return row;
}

function scanColumnForLabel_(sheet, label, scanCol, findLast) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) return -1;
  const vals = sheet.getRange(1, scanCol, lastRow, 1).getValues();
  let found = -1;
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === label) {
      found = i + 1; // 1-based
      if (!findLast) break;
    }
  }
  return found;
}

function pinRowMarker_(sheet, key, rowIndex) {
  sheet.createDeveloperMetadataFinder()
    .withLocationType(SpreadsheetApp.DeveloperMetadataLocationType.ROW)
    .withKey(key).find().forEach(m => m.remove());
  // Metadata requires an *entire* row — a bounded getRange(...maxColumns...) is
  // rejected as an "arbitrary range", so use A1 whole-row notation ("5:5").
  sheet.getRange(rowIndex + ":" + rowIndex).addDeveloperMetadata(key, DM_VISIBILITY);
}

/**
 * One-time (or anytime) primer: scans the Contact List the classic way and pins
 * metadata for every known landmark, so subsequent lookups are O(1). Safe to
 * re-run at any point — each pin is replaced. Exposed on the Custom Actions menu.
 */
function bootstrapContactListMarkers() {
  const [contactListSheet] = sheetsByName();

  [[MARKER_KEYS.EVENTS_ATTENDED, COL_CONSTANTS.EVENTS_ATTENDED],
   [MARKER_KEYS.EVENTS_RSVPD,    COL_CONSTANTS.EVENTS_RSVPD],
   [MARKER_KEYS.NEXT_STEPS,      COL_CONSTANTS.NEXT_STEPS]].forEach(([key, label]) => {
    Logger.log("col marker %s -> %s", label, findColMarker_(contactListSheet, key, label));
  });

  const totalRow = findRowMarker_(contactListSheet, MARKER_KEYS.TOTAL_RSVPD, COL_CONSTANTS.TOTAL_RSVPD, 2, true);
  Logger.log("row marker %s -> %s", COL_CONSTANTS.TOTAL_RSVPD, totalRow);
}
