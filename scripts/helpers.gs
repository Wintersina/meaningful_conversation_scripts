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
  return str.replace(/[\u2018\u2019\u201A\u201B\u2032\u2035]/g, "'"); // Replace curly and other apostrophes with standard single quote
}


function sheetsByName(){

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const scheduleSheet = ss.getSheetByName('Schedule');
  const eventbriteSheet = ss.getSheetByName("EventBrite Import");
  const contactListSheet = ss.getSheetByName("Contact List");

    if (!contactListSheet) {
    throw new Error("'Contact List' sheet not found.");
  }
  if (!eventbriteSheet) {
    throw new Error("'Schedule' sheet not found.");
  }

    if (!scheduleSheet) {
    throw new Error("'eventbriteSheet' sheet not found.");
  }

  return [contactListSheet, eventbriteSheet, scheduleSheet]
}