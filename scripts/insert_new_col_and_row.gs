function importNewEventsFromSchedule() {
  let [contactListSheet, _, scheduleSheet] = sheetsByName()

  const targetRow = 5; // Row to search
  const searchStringAttendedCol = "# Events Attended"

  // Get all values in the target row
  const rowValues = contactListSheet.getRange(targetRow, 1, 1, contactListSheet.getLastColumn()).getValues()[0];

  const AttendedColIndex = rowValues.indexOf(searchStringAttendedCol); // Add 1 because array index is 0-based
  if (AttendedColIndex == 0)
  {
    return;
  }

  // Fetch new titles from Schedule sheet and check if they are already in Contact List
  const scheduleTitles = scheduleSheet.getRange(1, 4, scheduleSheet.getLastRow(), 1).getValues().flat(); // Column D (1-based)
  const contactListTitles = contactListSheet.getRange(7, 1, 1, contactListSheet.getLastColumn()).getValues()[0];

  const normalizeString = (str) => str.replace(/[‘’‚‛′‵]/g, "'");
  const normalizedContactListTitles = contactListTitles.map(normalizeString);

  const newTitles = scheduleTitles.filter(title =>
    title && !["OFF", "TBD", "TOPIC"].includes(title.toUpperCase()) && !normalizedContactListTitles.includes(normalizeString(title))
  );

  // Skip if no new titles are found
  if (newTitles.length === 0) {
    Logger.log("No new titles found. Skipping column and row insertion.");
    return;
  }

  // Insert new column to the right of Column O
  const colOIndex = 15; // Column O (1-based index)
  const colPWidth = contactListSheet.getColumnWidth(colOIndex + 1); // Get width of Column P
  contactListSheet.insertColumnAfter(colOIndex);
  const newColIndex = colOIndex + 1;
  contactListSheet.setColumnWidth(newColIndex, colPWidth); // Set width of new column to match Column P

  // Insert VLOOKUP formula in Row 7 of the new column
  colLetter = columnToLetter(newColIndex)
  const row7 = 8;
  const vlookupFormula = `=VLOOKUP(${colLetter}${row7 - 1},'Event IDs'!$B:$C,2,FALSE)`;
  contactListSheet.getRange(row7, newColIndex).setFormula(vlookupFormula);

  //: Insert the first new title in Row 6 of the new column
  // todo update to a while there are new titles realy to be added
  // add all of them one at, do a for loop, the trickie part is
  // adding proper index for next rsvp/total index to be inserted.
  const row6 = 7;
  contactListSheet.getRange(row6, newColIndex).setValue(newTitles[0]);

  // Insert rows below "Total RSVP'd" and add formulas
  const data = contactListSheet.getRange(1, 2, contactListSheet.getLastRow(), 1).getValues().flat();
  const totalRSVPRowIndex = data.lastIndexOf("Total RSVP'd") + 1; // Convert to 1-based index

  if (totalRSVPRowIndex > 0) {
    const lastRSVPRowHeight = contactListSheet.getRowHeight(totalRSVPRowIndex); // Get height of "Total RSVP'd" row
    const insertRowIndex = totalRSVPRowIndex + 1;
    contactListSheet.insertRows(insertRowIndex, 1); // insert a row below the rsvp for title
    contactListSheet.insertRows(insertRowIndex + 5, 1); // insert the new total counter row for this


    contactListSheet.setRowHeight(insertRowIndex, lastRSVPRowHeight); // Set height of first inserted row
    contactListSheet.setRowHeight(insertRowIndex + 5, lastRSVPRowHeight); // Set height of second inserted row

    const grayRowRange1 = contactListSheet.getRange(insertRowIndex, 1, 1, contactListSheet.getLastColumn());
    const grayRowRange2 = contactListSheet.getRange(insertRowIndex + 5, 1, 1, contactListSheet.getLastColumn());
    grayRowRange1.setBackground('#D9D9D9');
    grayRowRange2.setBackground('#D9D9D9');

    // Use overflow text wrapping in inserted rows
    grayRowRange1.setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW);
    grayRowRange2.setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW);

    // Set horizontal alignment for the inserted rows
    grayRowRange1.setHorizontalAlignment('left');
    grayRowRange2.setHorizontalAlignment('left');

    // Insert "Total RSVP'd" in Column B of first gray row
    contactListSheet.getRange(insertRowIndex + 5, 2).setValue("Total RSVP'd");

    // Insert formula in first gray row (Columns P to AQ, not including AQ)
    const startColIndex = 16; // Column P starting index
    const endColIndex = AttendedColIndex + 2;
    for (let col = startColIndex; col < endColIndex; col++) {
      const columnLetter = columnToLetter(col)
      // we do + 4 here instead of five, as to not put the endin the formuala the same place as the formula itself.
      const formula = `=iferror(COUNTIF(${columnLetter}${insertRowIndex - 30}:${columnLetter}${insertRowIndex + 4},"RSVP'd: yes*"), "")`;
      const cell = contactListSheet.getRange(insertRowIndex + 5, col);

      // Insert the formula
      cell.setFormula(formula);

      // Ensure the cell is formatted as a number
      cell.setNumberFormat("0");
      cell.setHorizontalAlignment("center")
    }

    // Insert new title in Column B of second gray row
    contactListSheet.getRange(insertRowIndex, 2).setValue(newTitles[0]);

    // Add bottom border to "Total RSVP'd" row
    const totalRSVPRowRange = contactListSheet.getRange(insertRowIndex + 5, 2, 1, contactListSheet.getLastColumn() - 1); // Exclude last column
    totalRSVPRowRange.setBorder(false, false, true, false, false, false, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK);
  }
}

