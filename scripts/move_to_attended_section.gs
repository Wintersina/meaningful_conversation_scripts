
function copyAttendedToCorrectLocationAndPreserveRows() {
  let [sheet, _, _2] = sheetsByName()
  const data = sheet.getDataRange().getValues(); // Read all data from the sheet
  const formulas = sheet.getDataRange().getFormulas(); // Read all formulas (A1 notation)

  const colA = 0; // Column A (Index 0)
  const colB = 1; // Column B (Index 1)

  const targetRow = 5; // Row to search
  const searchStringRSVPCol = "# Events RSVP'd"; // String to search for
  const searchStringAttendedCol = "# Events Attended"

  // Get all values in the target row
  const rowValues = sheet.getRange(targetRow, 1, 1, sheet.getLastColumn()).getValues()[0];

  // Find the column index of the search string
  const RSVPColIndex = rowValues.indexOf(searchStringRSVPCol);
  const AttendedColIndex = rowValues.indexOf(searchStringAttendedCol);

  if (AttendedColIndex == 0)
  {
    return;
  }

  let RSVPIndex = null;
  let attended1PlusIndex = null;
  const rowsToCopy = [];

  // Find "Stop RSVP" in Column A
  for (let i = 0; i < data.length; i++) {
    if (data[i][colB] === "RSVP 2+") {
      RSVPIndex = i;
      break;
    }
  }
  if (RSVPIndex === null) {
    Logger.log("Stop RSVP not found in Column A.");
    return;
  }
  Logger.log("Stop RSVP found at row: " + (RSVPIndex + 1));

  // Find rows where Events Attended >= 1
  for (let i = RSVPIndex + 1; i < data.length; i++) {
    const eventsAttended = data[i][AttendedColIndex];
    if (eventsAttended >= 1) {
      rowsToCopy.push({ rowIndex: i, rowData: data[i], rowFormulas: formulas[i] });
      Logger.log(`Events Attended >= 1 at row ${i + 1}: ${data[i][colA]}`);
    }
  }

  // Find "ATTENDED 1+" in Column B
  for (let i = 0; i < data.length; i++) {
    if (data[i][colB] === "ATTENDED 1+") {
      attended1PlusIndex = i + 3;
      break;
    }
  }
  if (attended1PlusIndex === null) {
    Logger.log("ATTENDED 1+ not found in Column B.");
    return;
  }
  Logger.log("ATTENDED 1+ found at row: " + (attended1PlusIndex));

  // Insert copies of rows below "ATTENDED 1+" row, preserving formulas
  let insertionPoint = attended1PlusIndex; // Start inserting below "ATTENDED 1+"
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
}

// Function to adjust formula row references dynamically
// allows us to keep the original formual on the cell
function adjustFormulaForRow(formula, oldRow, newRow) {
  // Replace old row numbers with new row numbers in the formula
  return formula.replace(/\$?[A-Z]+(\d+)/g, (match, rowNum) => {
    const newRowNum = parseInt(rowNum, 10) + (newRow - oldRow);
    return match.replace(rowNum, newRowNum);
  });
}

