function sortRSVPRows() {
  /**
 * Sorts RSVP rows in the sheet based on event details while maintaining structure.
 *
 * This function:
 * 1. Retrieves columns A and B to locate:
 *    - "RSVP 2+" (starting point for sorting).
 *    - "Stop RSVP" (ending point for sorting).
 * 2. Identifies the range of RSVP rows to be sorted.
 * 3. Finds the column index for "Next Steps" to determine the sorting range.
 * 4. Sorts the RSVP rows based on multiple columns (from column P onward).
 * 5. Ensures the sorting is applied within the correct range without affecting other data.
 *
 * Key checks:
 * - Ensures "RSVP 2+" and "Stop RSVP" markers exist.
 * - Ensures sorting does not extend beyond valid row ranges.
 *
 * This script organizes RSVP data efficiently while preserving spreadsheet integrity.
 */
  Logger.log("starting sortRSVPRows");
  let [sheet, _, _2] = sheetsByName()
  const dataRange = sheet.getRange("A:B");  // Search columns A and B for "Stop RSVP" and "RSVP 2+"
  Logger.log("dataRange: " + dataRange);

  const data = dataRange.getValues();

  // Find the row where "RSVP 2+" is located
  let rsvpRow = data.findIndex(row => row[1] === COL_CONSTANTS.RSVP_2_PLUS);  // Look in column B
  if (rsvpRow === -1) return;  // If "RSVP 2+" not found, exit the function
  Logger.log("rsvpRow: " + rsvpRow);

  // Calculate the start row (2 rows below "RSVP 2+")
  let startRow = rsvpRow + 3;  // Adjust for the offset (13 is the starting row in the range)
  Logger.log("startRow: " + startRow);

  // Find the row where "Stop RSVP" is located
  // We'll use stop rsvp to know where the rsvp rows stop,
  // we can also use the "first" total rsvp'ed string from Columb B to get the same effect
  let stopRow = data.findIndex(row => row[0] === COL_CONSTANTS.STOP_RSVP);
  Logger.log("stopRow: " + stopRow);



  // Get all values in the target row
  const rowValues = sheet.getRange(ROW_NUMBERS.ROW_5, 1, 1, sheet.getLastColumn()).getValues()[0];

  // Find the column index of the search string
  const nextStepsIndex = rowValues.indexOf(COL_CONSTANTS.NEXT_STEPS) + 1; // Add 1 because array index is 0-based

  if (nextStepsIndex == 0){
    return
  }

  // Exit the function if stopRow not found or if our start row count is wrong
  // start row can be wrong, if we didnt find the correct starting area
  if (stopRow === -1 || stopRow + 13 < startRow) return;

  // Calculate the number of rows to sort (from startRow to one row below "Stop RSVP")
  const numRows = stopRow - startRow;  // Adjust to stop one row before "Stop RSVP"
  Logger.log("numRows: " + numRows);

  // Select the full range to sort (columns A to AF) for the calculated rows
  const range = sheet.getRange(startRow, 1, numRows, nextStepsIndex);  // XX columns (A to nextStepsIndex)
  Logger.log("range: " + range);



  // Create the sort specification based on columns P to XX (16 to XX) -- xx = any location the col many ends
  let sortSpecs = [];
  for (let i = 16; i <= nextStepsIndex; i++) {
    sortSpecs.push({ column: i, ascending: true });  // Adjust to `false` for descending order
  }
  Logger.log("sortSpecs: " + sortSpecs.length);

  range.sort(sortSpecs);
  Logger.log("ending sortRSVPRows");
}

