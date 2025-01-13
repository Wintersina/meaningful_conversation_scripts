
function moveRowsFromEventBriteImportToContactList() {
  // How it works
  // search for rows in our eventbrite import
  // group the list of events
  // sort the group to match the rows in the sheet, for preserving row index integeraty
  // add each grouped set to the respective rows
  // delete the moved over rows from the eventbrite import sheet bottom to top
  // profit
  let [contactListSheet, eventbriteSheet, _] = sheetsByName()

  const eventbriteData = eventbriteSheet.getRange(2, 1, eventbriteSheet.getLastRow() - 1, 12).getValues(); // Rows starting from 2
  const contactListData = contactListSheet.getRange(1, 2, contactListSheet.getLastRow()).getValues().flat(); // Column B
  const contactListTitles = contactListSheet.getRange(7, 1, 1, contactListSheet.getLastColumn()).getValues()[0]; // Row 7 (all columns)

  const targetRow = 5; // Row to search
  const searchStringRSVPCol = "# Events RSVP'd"; // String to search for
  const searchStringAttendedCol = "# Events Attended"

  // Get all values in the target row
  const rowValues = contactListSheet.getRange(targetRow, 1, 1, contactListSheet.getLastColumn()).getValues()[0];

  // Find the column index of the search string
  const RSVPColIndex = rowValues.indexOf(searchStringRSVPCol); // Add 1 because array index is 0-based
  const AttendedColIndex = rowValues.indexOf(searchStringAttendedCol); // Add 1 because array index is 0-based

  const columnLetter = columnToLetter(AttendedColIndex)

  if (RSVPColIndex === -1 || AttendedColIndex === -1) {
    return;
  }

  let rowsToDelete = []

  const groupedEvents = {};

  // Group rows by their cleaned event name
  eventbriteData.forEach((row, rowIndex) => {
    const originalEvent = row[10]; // Column K
    if (!originalEvent) return;

    // Clean the event name
    const cleanedEvent = normalizeString(originalEvent).replace(/\s*\(.*?\)\s*/g, "").trim();

    // Group rows by the cleaned event name
    if (!groupedEvents[cleanedEvent]) {
      groupedEvents[cleanedEvent] = [];
    }
    groupedEvents[cleanedEvent].push({ row, rowIndex });

  });

  const sortedGroupedEventsArray = Object.entries(groupedEvents).sort(([eventNameA], [eventNameB]) => {
    const indexA = contactListTitles.indexOf(eventNameA);
    const indexB = contactListTitles.indexOf(eventNameB);
    return indexA - indexB; // Sort by order in contactListTitles
  });

  // Convert the sorted array back into an object
  const sortedGroupedEvents = sortedGroupedEventsArray.reduce((acc, [eventName, rows]) => {
    acc[eventName] = rows;
    return acc;
  }, {});


  // let rows_have_shifted = 0
  Object.entries(sortedGroupedEvents).forEach(([eventName, rows]) => {
    // Find the matching event in Contact List
    const matchIndex = contactListData.indexOf(eventName) // + rows_have_shifted;
    if (matchIndex === -1) return;

    // Insert each row for the current event type
    rows.forEach(({row, rowIndex}) => {
      const valuesToCopy = row.slice(1, 12); // Copy B to L
      const insertRow = matchIndex + 3; // Match index is 0-based, +3 for row below match
      contactListSheet.insertRowBefore(insertRow);

      // Insert values starting at Column C (index 3)
      contactListSheet.getRange(insertRow, 3, 1, valuesToCopy.length).setValues([valuesToCopy]);

      // Add dynamic formulas
      const formula = `=IF(ISTEXT(C${insertRow}),CONCATENATE(C${insertRow}, " ", D${insertRow}),"")`;
      const formula_2 = `=IFERROR(COUNTIF(O${insertRow}:${columnLetter}${insertRow},"*Attended: yes"),"")`;
      const formula_3 = `=IFERROR(COUNTIF(O${insertRow}:${columnLetter}${insertRow},"RSVP'd: yes*"),"")`;

      contactListSheet.getRange(insertRow, 1).setFormula(formula); // A
      contactListSheet.getRange(insertRow, AttendedColIndex + 1).setFormula(formula_2); // AH
      contactListSheet.getRange(insertRow, RSVPColIndex + 1).setFormula(formula_3); // AI

      // Add formatted text in the corresponding column title
      const titleIndex = contactListTitles.indexOf(eventName);
      if (titleIndex !== -1) {
          const formattedText = `rsvp'd: yes\nattended: ?`;
          contactListSheet.getRange(insertRow, titleIndex + 1).setValue(formattedText);

          // Fill cells from the next column to AttendedColIndex with dashes
          const rangeStart = titleIndex + 2;  // the column immediately after the event column
          const rangeEnd = AttendedColIndex;  // the column index for 'Attended'

          // Fill the range with dashes, only if the range is valid
          if (rangeStart <= rangeEnd) {
            const dashRange = contactListSheet.getRange(insertRow, rangeStart, 1, rangeEnd - rangeStart + 1);
            const dashValues = dashRange.getValues();
            for (let i = 0; i < dashValues[0].length; i++) {
              dashValues[0][i] = '-';  // set each cell value to "-"
            }
            dashRange.setValues([dashValues[0]]);  // update the range with dashes
          }

      }
      rowsToDelete.push(rowIndex + 2);
    });
  })
    // fancy reverse
    rowsToDelete.sort((a, b) => b - a).forEach((rowIndex) => {
    eventbriteSheet.deleteRow(rowIndex);
  });
}
