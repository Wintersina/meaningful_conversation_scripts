function importNewEventsFromSchedule() {
/**
 * Imports new events from the Schedule sheet into the Contact List sheet.
 *
 * This function:
 * 1. Retrieves event titles and dates from the Schedule sheet.
 * 2. Checks if each scheduled event already exists in the Contact List sheet
 *    by comparing event titles and ensuring the date falls within a 7-day range.
 * 3. If an event is new, it:
 *    - Inserts a new column after Column O.
 *    - Copies the event date and title into the new column.
 *    - Adds VLOOKUP and RSVP tracking formulas.
 * 4. Adjusts the Contact List layout:
 *    - Inserts rows below the "Total RSVP’d" row to accommodate the new event.
 *    - Adds formulas to track RSVPs for the new event.
 *    - Applies consistent styling (gray background, text alignment, and borders).
 *
 *
 * This script helps maintain an up-to-date record of scheduled events in the
 * Contact List sheet, ensuring proper tracking and RSVP calculations.
 */


  let [contactListSheet, _, scheduleSheet] = sheetsByName()

  const targetRow = ROW_NUMBERS.ROW_5; // Row to search
  // Get all values in the target row
  const rowValues = contactListSheet.getRange(targetRow, 1, 1, contactListSheet.getLastColumn()).getValues()[0];

  const AttendedColIndex = rowValues.indexOf(COL_CONSTANTS.EVENTS_ATTENDED); // Add 1 because array index is 0-based
  if (AttendedColIndex == 0)
  {
    return;
  }

  const scheduleData = scheduleSheet.getRange(1, 3, scheduleSheet.getLastRow(), 2).getValues(); // Columns C & D


  // Fetch new titles from Schedule sheet and check if they are already in Contact List
  const contactListTitles = contactListSheet.getRange(7, 1, 1, contactListSheet.getLastColumn()).getValues()[0];
  const contactListDates = contactListSheet.getRange(6, 1, 1, contactListSheet.getLastColumn()).getValues()[0];

  const localNormalizeString = (str) => normalizeString(str).toUpperCase();

  const isWithinWeek = (scheduleDateStr, contactDateStr) => {
    if (!scheduleDateStr || !contactDateStr) return false;

    try {
      let scheduleDate = parseDate(scheduleDateStr); // Convert schedule format
      let contactDate = parseDate(contactDateStr); // Convert contact new format

      let diffDays = Math.abs((scheduleDate - contactDate) / (1000 * 60 * 60 * 24)); // Difference in days
      return diffDays >= -7 && diffDays <= 7;; // Check if within a 7-day range
    } catch (e) {
      Logger.log("Error parsing dates: " + e.message);
      return false;
    }
  };

  // this is needed because the date format in sch sheet is different
  // then the date field in contact list we have to parse this...
    const parseDate = (dateStr) => {
      if (!dateStr) return null;

      // If it's already a Date object, return a new Date without the time part
      if (dateStr instanceof Date) {
        return new Date(dateStr.getFullYear(), dateStr.getMonth(), dateStr.getDate());
      }

      return null; // If it's not a Date, return null
    };


  let newEvents = scheduleData.filter(([date, title]) => {
    if (!title || SCHEDULE_SHEET_CONSTANTS.SKIP_WORDS.includes(title.toUpperCase())) return false;

    let normalizedTitle = localNormalizeString(title);
    let matchFound = contactListTitles.some((existingTitle, index) =>
      normalizedTitle === localNormalizeString(existingTitle) &&
      isWithinWeek(date, contactListDates[index])
    );

    return !matchFound; // Keep only new events
  });

  // skip when we dont find any events
  if (newEvents.length === 0) {
    Logger.log("No new unique events found. Skipping.");
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

  contactListSheet.getRange(ROW_NUMBERS.ROW_8, newColIndex).setValue(newEvents[0][0]); // put  Date
  contactListSheet.getRange(ROW_NUMBERS.ROW_7, newColIndex).setValue(newEvents[0][1]); //  put Event Name

  Logger.log("Added new event: " + newEvents[0][1]);

  contactListSheet.getRange(ROW_NUMBERS.ROW_8, newColIndex).setFormula(FORMULAS.V_LOOKUP(colLetter));
  contactListSheet.getRange(ROW_NUMBERS.ROW_9 , newColIndex).setFormula(FORMULAS.TTL_RSVP(colLetter)).setFontSize(9)
  contactListSheet.getRange(ROW_NUMBERS.ROW_10, newColIndex).setFormula(FORMULAS.TTL_ATTND(colLetter)).setFontSize(9)


  //: Insert the first new title in Row 6 of the new column
  // todo update to a while there are new titles realy to be added
  // add all of them one at, do a for loop, the trickie part is
  // adding proper index for next rsvp/total index to be inserted.

  // Insert rows below "Total RSVP'd" and add formulas
  const data = contactListSheet.getRange(1, 2, contactListSheet.getLastRow(), 1).getValues().flat();
  const totalRSVPRowIndex = data.lastIndexOf(COL_CONSTANTS.TOTAL_RSVPD) + 1; // Convert to 1-based index

  if (totalRSVPRowIndex > 0) {
    const lastRSVPRowHeight = contactListSheet.getRowHeight(totalRSVPRowIndex); // Get height of "Total RSVP'd" row
    const insertRowIndex = totalRSVPRowIndex + 1;
    contactListSheet.insertRows(insertRowIndex, 1); // insert a row below the rsvp for title
    contactListSheet.insertRows(insertRowIndex + 5, 1); // insert the new total counter row for this


    contactListSheet.setRowHeight(insertRowIndex, lastRSVPRowHeight); // Set height of first inserted row
    contactListSheet.setRowHeight(insertRowIndex + 5, lastRSVPRowHeight); // Set height of second inserted row

    const grayRowRange1 = contactListSheet.getRange(insertRowIndex, 1, 1, contactListSheet.getLastColumn());
    const grayRowRange2 = contactListSheet.getRange(insertRowIndex + 5, 1, 1, contactListSheet.getLastColumn());
    grayRowRange1.setBackground(UI_CONSTANTS.GRAY_BACKGROUND);
    grayRowRange2.setBackground(UI_CONSTANTS.GRAY_BACKGROUND);

    // Use overflow text wrapping in inserted rows
    grayRowRange1.setWrapStrategy(UI_CONSTANTS.WRAP_STRATEGY);
    grayRowRange2.setWrapStrategy(UI_CONSTANTS.WRAP_STRATEGY);

    // Set horizontal alignment for the inserted rows
    grayRowRange1.setHorizontalAlignment('left');
    grayRowRange2.setHorizontalAlignment('left');

    // Insert "Total RSVP'd" in Column B of first gray row
    contactListSheet.getRange(insertRowIndex + 5, 2).setValue(COL_CONSTANTS.TOTAL_RSVPD);

    // Insert formula in first gray row (Columns P to AQ, not including AQ)
    const startColIndex = 16; // Column P starting index
    const endColIndex = AttendedColIndex + 2;
    for (let col = startColIndex; col < endColIndex; col++) {
      const columnLetter = columnToLetter(col)
      // we do + 4 here instead of five, as to not put the endin the formuala the same place as the formula itself.
      // TODO MAKE THIS A CONSTANT FORMULA, BECAUSE OF RANGE ITS NOT EASY
      const formula = `=iferror(COUNTIF(${columnLetter}${insertRowIndex - 30}:${columnLetter}${insertRowIndex + 4},"RSVP'd: yes*"), "")`;
      const cell = contactListSheet.getRange(insertRowIndex + 5, col);

      // Insert the formula
      cell.setFormula(formula);

      // Ensure the cell is formatted as a number
      cell.setNumberFormat("0");
      cell.setHorizontalAlignment(UI_CONSTANTS.ALIGNMENT_CENTER)
    }

    // Insert new title in Column B of second gray row
    contactListSheet.getRange(insertRowIndex, 2).setValue(newEvents[0][1]);

    // Add bottom border to "Total RSVP'd" row
    const totalRSVPRowRange = contactListSheet.getRange(insertRowIndex + 5, 2, 1, contactListSheet.getLastColumn() - 1); // Exclude last column
    totalRSVPRowRange.setBorder(false, false, true, false, false, false, UI_CONSTANTS.BORDER_COLOR, UI_CONSTANTS.BORDER_STYLE);
  }
}

