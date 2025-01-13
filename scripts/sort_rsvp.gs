function sortRSVPRows() {
  let [sheet, _, _2] = sheetsByName()
  const dataRange = sheet.getRange("A:B");  // Search columns A and B for "Stop RSVP" and "RSVP 2+"
  Logger.log("dataRange: " + dataRange);

  const data = dataRange.getValues();  // Get values from columns A and B

  // Logger.log("data: " + data);

  // Find the row where "RSVP 2+" is located
  let rsvpRow = data.findIndex(row => row[1] === "RSVP 2+");  // Look in column B
  if (rsvpRow === -1) return;  // If "RSVP 2+" not found, exit the function
  Logger.log("rsvpRow: " + rsvpRow);

  // Calculate the start row (2 rows below "RSVP 2+")
  let startRow = rsvpRow + 3;  // Adjust for the offset (13 is the starting row in the range)
  Logger.log("startRow: " + startRow);

  // Find the row where "Stop RSVP" is located
  let stopRow = data.findIndex(row => row[0] === "Stop RSVP");  // Look in column A
  Logger.log("stopRow: " + stopRow);

  const targetRow = 5; // Row to search
  const searchStringRSVPCol = "# Events RSVP'd"; // String to search for

  // Get all values in the target row
  const rowValues = sheet.getRange(targetRow, 1, 1, sheet.getLastColumn()).getValues()[0];

  // Find the column index of the search string
  const RSVPColIndex = rowValues.indexOf(searchStringRSVPCol) + 1; // Add 1 because array index is 0-based

  if (RSVPColIndex == 0){
    return
  }

  if (stopRow === -1 || stopRow + 13 < startRow) return;

  // Calculate the number of rows to sort (from startRow to one row below "Stop RSVP")
  const numRows = stopRow - startRow;  // Adjust to stop one row before "Stop RSVP"
  Logger.log("numRows: " + numRows);

  // Select the full range to sort (columns A to AF) for the calculated rows
  const range = sheet.getRange(startRow, 1, numRows, RSVPColIndex);  // XX columns (A to RSVPColIndex)
  Logger.log("range: " + range);



  // Create the sort specification based on columns P to XX (16 to XX) -- xx = any location the col many ends
  let sortSpecs = [];
  for (let i = 16; i <= RSVPColIndex; i++) {
    sortSpecs.push({ column: i, ascending: true });  // Adjust to `false` for descending order
  }
  Logger.log("sortSpecs: " + sortSpecs.length);

  // Sort the full range based on the specified columns
  range.sort(sortSpecs);
}

