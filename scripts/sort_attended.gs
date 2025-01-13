function sortAttendedRows() {
  let [sheet, _, _2] = sheetsByName()
  // Search column B from row 13 onwards
  const dataRange = sheet.getRange("B13:B");
  // Get all values from the range
  const data = dataRange.getValues();

  const targetRow = 5; // Row to search
  const searchStringRSVPCol = "# Events RSVP'd"; // String to search for

  // Get all values in the target row
  const rowValues = sheet.getRange(targetRow, 1, 1, sheet.getLastColumn()).getValues()[0];

  // Find the column index of the search string
  const RSVPColIndex = rowValues.indexOf(searchStringRSVPCol) + 1; // Add 1 because array index is 0-based

  if (RSVPColIndex == 0){
    return;
  }

  // Find the row where "RSVP 2+" is located
  // Adjust for offset starting at row 13
  let endRow = data.findIndex(row => row[0] === "RSVP 2+") + 13;
  // If not found, stop the function, because we dont wanna sort dangerously
  if (endRow === 12) return;

  // Calculate the number of rows to sort (from row 13 to just before "RSVP 2+")
  const numRows = endRow - 13;

  // Select the range from column P to AF for the dynamic number of rows
  const range = sheet.getRange(13, 1, numRows, RSVPColIndex);  // XX amount columns (A to XX)

  // Create a sort specification for columns P to AF
  let sortSpecs = [];
  for (let i = 16; i <= RSVPColIndex; i++) {
    sortSpecs.push({ column: i, ascending: true });
  }


  // Sort the range based on the specified order
  range.sort(sortSpecs);
}


