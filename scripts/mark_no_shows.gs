/**
 * Marks "rsvp'd: yes / attended: ?" as "rsvp'd: yes / attended: no"
 * for event columns whose date (Row 6) is more than 1 day in the past.
 *
 * Only processes events within the last 2 months to avoid wasted effort.
 * Starts scanning contact rows at row 12.
 */
function markNoShows() {
  Logger.log("starting markNoShows");

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.CONTACT_LIST);
  if (!sheet) throw new Error(ERROR_MESSAGES.CONTACT_LIST_NOT_FOUND);

  var lastCol = sheet.getLastColumn();
  var lastRow = sheet.getLastRow();
  var startCol = HELPER_CONSTANTS.EVENT_NAMES_START_COL; // column O (15)

  if (lastCol < startCol || lastRow < ROW_NUMBERS.ROW_12) {
    Logger.log("Nothing to process");
    return;
  }

  // Read all event dates from Row 6 (columns O onward)
  var datesRange = sheet.getRange(ROW_NUMBERS.ROW_6, startCol, 1, lastCol - startCol + 1);
  var dates = datesRange.getValues()[0];

  var now = new Date();
  var oneDayAgo = new Date(now.getTime() - (1 * 24 * 60 * 60 * 1000));
  var twoMonthsAgo = new Date(now.getFullYear(), now.getMonth() - 2, now.getDate());

  // Find which columns need processing
  var colsToProcess = [];
  for (var i = 0; i < dates.length; i++) {
    var eventDate = dates[i];
    if (!(eventDate instanceof Date) || isNaN(eventDate.getTime())) continue;

    // Skip if event hasn't passed yet (less than 1 day ago)
    if (eventDate >= oneDayAgo) continue;

    // Skip if older than 2 months
    if (eventDate < twoMonthsAgo) continue;

    colsToProcess.push(startCol + i); // 1-based column index
  }

  if (colsToProcess.length === 0) {
    Logger.log("No event columns need updating");
    return;
  }

  Logger.log("Processing " + colsToProcess.length + " event columns");

  var dataRowStart = ROW_NUMBERS.ROW_12;
  var numRows = lastRow - dataRowStart + 1;

  // Process each qualifying column
  var totalUpdated = 0;
  for (var c = 0; c < colsToProcess.length; c++) {
    var col = colsToProcess[c];
    var range = sheet.getRange(dataRowStart, col, numRows, 1);
    var values = range.getValues();
    var updated = false;

    for (var r = 0; r < values.length; r++) {
      if (values[r][0] === RSVP_DROP_DOWN_CONSTANTS.YES_ATTENDED) {
        values[r][0] = RSVP_DROP_DOWN_CONSTANTS.YES_ATTENDED_NO;
        updated = true;
        totalUpdated++;
      }
    }

    if (updated) {
      range.setValues(values);
      Logger.log("Updated column " + columnToLetter(col));
    }
  }

  Logger.log("markNoShows complete. " + totalUpdated + " cells updated across " + colsToProcess.length + " columns");
}
