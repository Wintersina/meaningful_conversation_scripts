const SHEET_NAMES = {
  SCHEDULE: "Schedule",
  EVENTBRITE: "EventBrite Import",
  CONTACT_LIST: "Contact List"
};

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
  // TOTAL_ATTENDED: "# Events RSVP'd",
  STOP_RSVP: "Stop RSVP",
  RSVP_2_PLUS: "RSVP 2+",
  ATTENDED_PLUS_ONE:"ATTENDED 1+"
}

const RSVP_DROP_DOWN_CONSTANTS = {
  YES_ATTENDED: "rsvp'd: yes\nattended: ?",
  MAYBE_ATTENDED: "rsvp'd: maybe\nattended: ?"
}

const FORMULAS = {
  CONCATENATE_NAME: (row) => `=IF(ISTEXT(C${row}),CONCATENATE(C${row}, " ", D${row}),"")`,
  COUNT_ATTENDED: (row, colLetter) => `=IFERROR(COUNTIF(O${row}:${colLetter}${row},"*Attended: yes"),"")`,
  COUNT_RSVP: (row, colLetter) => `=IFERROR(COUNTIF(O${row}:${colLetter}${row},"RSVP'd: yes*"),"")`,
  V_LOOKUP:(colLetter) => `=VLOOKUP(${colLetter}${ROW_NUMBERS.ROW_7},'Event IDs'!$B:$C,2,FALSE)`,
  TTL_RSVP:(colLetter) => `="TTL RSVP =" & COUNTIF(${colLetter}13:${colLetter}2008, "*RSVP'D: yes*")`,
  TTL_ATTND:(colLetter) => `="TTL ATTND =" & COUNTIF(${colLetter}13:${colLetter}2009, "*ATTENDED: yes*")`
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
  ROW_11 : 11,
  ROW_12 : 12,

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
