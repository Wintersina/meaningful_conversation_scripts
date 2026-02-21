function mergeRowsByKeyPreserveAllFormulas() {

/**
 * Merges duplicate rows in the sheet while preserving all formulas.
 * this is probably the slowest method we have
 * This function:
 * 1. Retrieves all data and formulas from the sheet.
 * 2. Identifies duplicate rows based on the key in Column A.
 * 3. Merges values from duplicate rows into a single row while:
 *    - Preserving formulas.
 *    - Concatenating unique values.
 *    - Skipping specific columns that should not be merged.
 * 4. Updates the sheet with merged rows in a batch operation for efficiency.
 * 5. Deletes duplicate rows in reverse order to prevent row shifting issues.
 *
 * Performance Considerations:
 * - This method is slow for large sheets, especially due to calls like `getLastColumn()`,
 *   which are expensive operations.
 * - Performance could be improved by defining a fixed column index range if known.
 *
 * Key Implementation Details:
 * - `skipCols`: Defines columns that should not be merged.
 * - `mergedData`: Stores merged values and formulas in memory before batch updating.
 * - `rowsToDelete`: Keeps track of duplicate rows to be deleted efficiently.
 *
 * This script ensures that duplicate rows are merged efficiently while preserving
 * the sheet's structure and formulas.
 */
  Logger.log("starting mergeRowsByKeyPreserveAllFormulas");
  let [sheet, _, _2] = sheetsByName()

  var range = sheet.getDataRange();
  var data = range.getValues();
  var formulas = range.getFormulas();

  var mergedData = {};
  var rowsToDelete = [];



  // Get all values in the target row
  const rowValues = sheet.getRange(ROW_NUMBERS.ROW_5, 1, 1, sheet.getLastColumn()).getValues()[0];

  // Find the column index of the search string
  const RSVPColIndex = rowValues.indexOf(COL_CONSTANTS.EVENTS_RSVPD); // Add 1 because array index is 0-based
  const AttendedColIndex = rowValues.indexOf(COL_CONSTANTS.EVENTS_ATTENDED); // Add 1 because array index is 0-based


  if (RSVPColIndex == 0)
  {
    return;
  }


  if (AttendedColIndex == 0)
  {
    return;
  }

  // this is not the cleanest way to ignore col's we wanna merge
  // however it gets the job done.
  // a sligtly better approach would be, to define what each index and column is,
  // create a allow list of sorts that can be managed as a helper
  // use the allow list for the filter
  var skipCols = [0, 1, 2, 3,9,11, RSVPColIndex, AttendedColIndex, RSVPColIndex +1]; // A=0, B=1, C=2, D=3, AG=32, AH=33, AI=34 J=9


  // Process rows and merge duplicates in memory
  for (var i = 1; i < data.length; i++) { // Skip header row
    var key = normalizeByStrippingWhiteSpaceAtTheEnd(data[i][0]) // Use Column A as the key
    if (!key) continue; // Skip empty keys

    if (!mergedData[key]) {
      // Store the first occurrence of the key
      mergedData[key] = {
        values: data[i].slice(), // Clone row values
        formulas: formulas[i].slice(), // Clone row formulas
        rowIndex: i + 1 // Track 1-based row index
      };
    } else {
      // Merge the current row into the existing row
      for (var col = 0; col < data[i].length; col++) {
        if (skipCols.includes(col)) continue; // Skip specified columns

        var existingValue = mergedData[key].values[col];
        var newValue = data[i][col];
        var existingFormula = mergedData[key].formulas[col];
        var newFormula = formulas[i][col];

        // Preserve formulas
        if (!existingFormula && newFormula) {
          mergedData[key].formulas[col] = newFormula;
        }

        // Merge non-empty values only if not skipped
        if (!existingValue || existingValue === "-") {
          if (newValue && newValue !== "-") {
            mergedData[key].values[col] = newValue;


        } else if (existingValue !== newValue && newValue && newValue !== "-") {
          mergedData[key].values[col] = existingValue + ", " + newValue;
        }
        }

      }
      // Mark duplicate row for deletion
      rowsToDelete.push(i + 1);
    }
  }

  // Batch update only the merged rows
  Object.values(mergedData).forEach(function(entry) {
    var rowIndex = entry.rowIndex;
    var rowValues = entry.values;
    var rowFormulas = entry.formulas;

    // Write the entire row in one call
    var rowRange = sheet.getRange(rowIndex, 1, 1, rowValues.length);
    rowRange.setValues([rowValues]);

    // Apply formulas where necessary
    rowFormulas.forEach(function(formula, colIndex) {
      if (formula) {
        sheet.getRange(rowIndex, colIndex + 1).setFormula(formula);
      }
    });
  });

  // Delete duplicate rows in reverse order to avoid shifting
  rowsToDelete.sort(function(a, b) { return b - a; }); // Sort descending
  rowsToDelete.forEach(function(rowIndex) {
    sheet.deleteRow(rowIndex);
  });

  Logger.log("ending mergeRowsByKeyPreserveAllFormulas");

}
