
function moveRowsFromEventBriteImportToContactList() {
/**
 * Moves rows from the Eventbrite Import sheet to the Contact List sheet.
 *
 * This function:
 * 1. Retrieves data from the Eventbrite Import sheet and Contact List sheet.
 * 2. Groups imported Eventbrite rows by event name after normalizing them.
 * 3. Sorts grouped events to match the order in the Contact List sheet to maintain row integrity.
 * 4. For each grouped event:
 *    - Finds the corresponding row in the Contact List sheet.
 *    - Inserts new rows under the matching event title.
 *    - Copies attendee data into the respective columns.
 *    - Adds dynamic formulas for tracking attendance and RSVPs.
 *    - Fills empty columns between event and "Attended" with dashes.
 * 5. Deletes the moved rows from the Eventbrite Import sheet (bottom to top to preserve indexing).
 *
 * `columnToLetter(index)`: Converts a column index to its corresponding letter.
 *
 * This script ensures a structured transfer of event data while maintaining
 * formatting and RSVP tracking in the Contact List sheet.
 */

  let [contactListSheet, eventbriteSheet, _] = sheetsByName()

const lastRow = eventbriteSheet.getLastRow();
if (lastRow <= 1) {
  return; // Exit script if no data is found
}

  const eventbriteData = eventbriteSheet.getRange(2, 1, lastRow - 1, 12).getValues(); // Rows starting from 2
  const contactListData = contactListSheet.getRange(1, 2, contactListSheet.getLastRow()).getValues().flat(); // Column B

  const contactListTitles = contactListSheet.getRange(7, 1, 1, contactListSheet.getLastColumn()).getValues()[0]; // Row 7 (all columns)

  const rowValues = contactListSheet.getRange(ROW_NUMBERS.ROW_5, 1, 1, contactListSheet.getLastColumn()).getValues()[0];

  const RSVPColIndex = rowValues.indexOf(COL_CONSTANTS.EVENTS_RSVPD); // Add 1 because array index is 0-based
  const AttendedColIndex = rowValues.indexOf(COL_CONSTANTS.EVENTS_ATTENDED); // Add 1 because array index is 0-based

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
    const matchIndex = contactListData.lastIndexOf(eventName.trim()) // + rows_have_shifted;
    // todo, make trim a proper helper and pass every name like thing to it
    if (matchIndex === -1) return;

    // Insert each row for the current event type
    rows.forEach(({row, rowIndex}) => {
      const valuesToCopy = row.slice(1, 12); // Copy B to L
      const insertRow = matchIndex + 3; // Match index is 0-based, +3 for row below match
      contactListSheet.insertRowBefore(insertRow);

      // Insert values starting at Column C (index 3)
      contactListSheet.getRange(insertRow, 3, 1, valuesToCopy.length).setValues([valuesToCopy]);

      // set row formatting
      const newRowRange = contactListSheet.getRange(insertRow, 1, 1, contactListSheet.getLastColumn());

      newRowRange.setFontSize(UI_CONSTANTS.FONT_SIZE);
      newRowRange.setFontFamily(UI_CONSTANTS.FONT_STYLE);
      newRowRange.setHorizontalAlignment(UI_CONSTANTS.ALIGNMENT_CENTER);

      // Add dynamic formulas
      contactListSheet.getRange(insertRow, 1).setFormula(FORMULAS.CONCATENATE_NAME(insertRow)); // A
      contactListSheet.getRange(insertRow, AttendedColIndex + 1).setFormula(FORMULAS.COUNT_ATTENDED(insertRow, columnLetter));
      contactListSheet.getRange(insertRow, RSVPColIndex + 1).setFormula(FORMULAS.COUNT_RSVP(insertRow, columnLetter));

      // Add formatted text in the corresponding column title
      const titleIndex = contactListTitles.indexOf(eventName.trim());
      if (titleIndex !== -1) {

          let formula_to_use = RSVP_DROP_DOWN_CONSTANTS.YES_ATTENDED

          // todo set maybe formula for facebook users who are maybe
          // if (facebook) {
          //   formula_to_use = RSVP_DROP_DOWN_CONSTANTS.MAYBE_ATTENDED
          // }

          contactListSheet.getRange(insertRow, titleIndex + 1).setValue(formula_to_use);

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
