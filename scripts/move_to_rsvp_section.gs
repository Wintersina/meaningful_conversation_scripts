
function copyRSVPToCorrectLocationAndPreserveRows() {
/**
 * Copies RSVP rows with at least 2+ RSVPs to the correct location while preserving formulas.
 *
 * This function:
 * 1. Retrieves all data and formulas from the sheet.
 * 2. Finds the "RSVP 2+" and "Stop RSVP" markers in the sheet to determine where data should be copied.
 * 3. Identifies rows where "Events RSVP’d" is 2 or more.
 * 4. Copies these rows below the "RSVP 2+" marker while:
 *    - Inserting new rows.
 *    - Preserving original formulas.
 *    - Adjusting row references dynamically in formulas.
 * 5. Calls `mergeRowsByKeyPreserveAllFormulas()` to clean up and merge duplicates after inserting.
 *
 * Key functions:
 * - `adjustFormulaForRow(formula, oldRow, newRow)`: Adjusts row references dynamically when copying formulas.
 * - `mergeRowsByKeyPreserveAllFormulas()`: Merges duplicate entries while keeping formulas intact.
 *
 * This script ensures proper RSVP tracking by moving relevant rows while maintaining
 * sheet integrity and formulas.
 */

  Logger.log("starting copyRSVPToCorrectLocationAndPreserveRows");

  let [sheet, _, _2] = sheetsByName()
  const data = sheet.getDataRange().getValues();
  const formulas = sheet.getDataRange().getFormulas();

  const colA = 0; // Column A (Index 0)
  const colB = 1; // Column B (Index 1)


  // Get all values in the target row
  const rowValues = sheet.getRange(ROW_NUMBERS.ROW_5, 1, 1, sheet.getLastColumn()).getValues()[0];

  // Find the column index of the search string
  const RSVPColIndex = rowValues.indexOf(COL_CONSTANTS.EVENTS_RSVPD);
  if (RSVPColIndex == 0)
  {
    return;
  }
  let RSVPIndex = null;

  const rowsToCopy = [];

  // Find "Stop RSVP" in Column A
  for (let i = 0; i < data.length; i++) {
    if (data[i][colB] === COL_CONSTANTS.RSVP_2_PLUS) {
      RSVPIndex = i;
      break;
    }
  }
  if (RSVPIndex === null) {
    Logger.log("Stop RSVP not found in Column B.");
    return;
  }
  Logger.log("RSVP 2+ found at row: " + (RSVPIndex + 1));

  for (let i = 0; i < data.length; i++) {
    if (data[i][colA] === COL_CONSTANTS.STOP_RSVP) {
      StopRSVPIndex = i;
      break;
    }
  }
  if (StopRSVPIndex === null) {
    Logger.log("Stop RSVP not found in Column A.");
    return;
  }
  Logger.log("Stop RSVP found at row: " + (StopRSVPIndex + 1));


  // Find rows where Events Attended >= 1
  for (let i = StopRSVPIndex + 5; i < data.length; i++) {
    const eventsRSVP = data[i][RSVPColIndex];
    if (eventsRSVP >= 2) {
      rowsToCopy.push({ rowIndex: i, rowData: data[i], rowFormulas: formulas[i] });
      Logger.log(`Events Attended >= 1 at row ${i + 1}: ${data[i][colA]}`);
    }
  }

  let RSVP2Start = RSVPIndex + 3

  // Insert copies of rows below "ATTENDED 1+" row, preserving formulas
  let insertionPoint = RSVP2Start; // Start inserting below "ATTENDED 1+"
  rowsToCopy.forEach(row => {
    sheet.insertRowsBefore(insertionPoint, 1); // Insert a blank row

    // Copy values and formulas into the new row
    for (let col = 0; col < row.rowData.length; col++) {
      const cell = sheet.getRange(insertionPoint, col + 1);
      if (row.rowFormulas[col]) {
        const adjustedFormula = adjustFormulaForRow(row.rowFormulas[col], row.rowIndex + 1, insertionPoint);
        cell.setFormula(adjustedFormula); // Set A1 formula if present
      } else {
        cell.setValue(row.rowData[col]); // Set value if no formula
      }
    }
    Logger.log(`Inserted copy of row ${row.rowIndex + 1} at row ${insertionPoint}`);
    insertionPoint++; // Update insertion point
  });

  // now that the row has been copied to correct location, we merge the duplicates
  mergeRowsByKeyPreserveAllFormulas()
  Logger.log("ending copyRSVPToCorrectLocationAndPreserveRows");
}

// Function to adjust formula row references dynamically
// allows us to keep the original formulas in place, in respective cell
function adjustFormulaForRow(formula, oldRow, newRow) {
  // Replace old row numbers with new row numbers in the formula
  return formula.replace(/\$?[A-Z]+(\d+)/g, (match, rowNum) => {
    const newRowNum = parseInt(rowNum, 10) + (newRow - oldRow);
    return match.replace(rowNum, newRowNum);
  });
}

