function mergeRowsByKeyPreserveAllFormulas() {
  let [sheet, _, _2] = sheetsByName()

  var range = sheet.getDataRange();
  var data = range.getValues();
  var formulas = range.getFormulas();

  var mergedData = {};
  var rowsToDelete = [];

  // Columns to skip during the merge (0-based index)

  const targetRow = 5; // Row to search
  const searchStringRSVPCol = "# Events RSVP'd"; // String to search for
  const searchStringAttendedCol = "# Events Attended"

  // Get all values in the target row
  const rowValues = sheet.getRange(targetRow, 1, 1, sheet.getLastColumn()).getValues()[0];

  // Find the column index of the search string
  const RSVPColIndex = rowValues.indexOf(searchStringRSVPCol); // Add 1 because array index is 0-based
  const AttendedColIndex = rowValues.indexOf(searchStringAttendedCol); // Add 1 because array index is 0-based


  if (RSVPColIndex == 0)
  {
    return;
  }


  if (AttendedColIndex == 0)
  {
    return;
  }

  var skipCols = [0, 1, 2, 3,9,11, RSVPColIndex, AttendedColIndex, RSVPColIndex +1]; // A=0, B=1, C=2, D=3, AG=32, AH=33, AI=34 J=9


  // Process rows and merge duplicates in memory
  for (var i = 1; i < data.length; i++) { // Skip header row
    var key = data[i][0] ? data[i][0].toLowerCase() : null; // Use Column A as the key
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
        if (!existingValue) {
          mergedData[key].values[col] = newValue; // Replace empty cell or a cell with '-' in it
        } else if (existingValue !== newValue && newValue) {
          if (existingValue === "-")
          {
            mergedData[key].values[col] = newValue
          }
          else if (newValue === "-"){
            mergedData[key].values[col] = existingValue
          }
          else{
          mergedData[key].values[col] = existingValue + ", " + newValue; // Concatenate
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
}
