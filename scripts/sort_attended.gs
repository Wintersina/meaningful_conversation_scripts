function sortAttendedRows() {
/**
 * Sorts "Attended" rows in the sheet based on event details while preserving structure.
 *
 * This function:
 * 1. Retrieves data from column B starting at row 13 to locate:
 *    - "RSVP 2+" (which marks the end of the sorting range).
 * 2. Identifies the column index of "Next Steps" to determine the sorting width.
 * 3. Defines the sorting range:
 *    - Starts at row 13.
 *    - Ends just before "RSVP 2+".
 *    - Includes all relevant columns from "P" to "Next Steps".
 * 4. Sorts the selected rows based on multiple columns in ascending order.
 *
 * Key checks:
 * - Ensures "RSVP 2+" exists before sorting to prevent errors.
 * - Ensures sorting does not extend beyond the correct row range.
 * - Adjusts for dynamic columns using `nextStepsIndex`.
 *
 * This script helps maintain an organized event attendance list while keeping
 * data integrity intact.
 */
  Logger.log("starting sortAttendedRows");


  let [sheet, _, _2] = sheetsByName()
  // Search column B from row 13 onwards
  const dataRange = sheet.getRange("B13:B");

  const data = dataRange.getValues();

  const targetRow = ROW_NUMBERS.ROW_5

  const rowValues = sheet.getRange(targetRow, 1, 1, sheet.getLastColumn()).getValues()[0];

  // Find the column index of the search string
  // + 6 if we wanna include notes and other columns that we normally not include in ours
  // we can also include the proper search col, of the last actual column by name and set that as the index
  const nextStepsIndex = rowValues.indexOf(COL_CONSTANTS.NEXT_STEPS) + 1; // Add 1 because array index is 0-based

  if (nextStepsIndex == 0){
    return;
  }

  // Find the row where "RSVP 2+" is located
  // Adjust for offset starting at row 13
  let endRow = data.findIndex(row => row[0] === COL_CONSTANTS.RSVP_2_PLUS) + 13;
  // If not found, stop the function, because we dont wanna sort dangerously
  if (endRow === 12) return;

  // Calculate the number of rows to sort (from row 13 to just before "RSVP 2+")
  const numRows = endRow - 13;

  // Select the range from column P to XX for the dynamic number of rows
  const range = sheet.getRange(13, 1, numRows, nextStepsIndex);  // XX amount columns (A to XX)

  // Create a sort specification for columns O to XX
  // insert sort criteria
  // o == 15
  let sortSpecs = [];
  for (let i = HELPER_CONSTANTS.EVENT_NAMES_START_COL; i <= nextStepsIndex; i++) {
    sortSpecs.push({ column: i, ascending: true });
  }


  // Sort the range based on the specified order
  range.sort(sortSpecs);
  Logger.log("ending sortAttendedRows");
}


