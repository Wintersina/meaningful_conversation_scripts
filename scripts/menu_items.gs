function onOpen() {
  /**
 * Adds a custom menu to the Google Sheets UI when the spreadsheet is opened.
 *
 * This function:
 * 1. Retrieves the active sheet.
 * 2. If the sheet name is "Contact List":
 *    - Creates a "Custom Actions" menu in the UI.
 *    - Adds menu items that trigger various automation functions, including:
 *      - Sorting attended and RSVP rows.
 *      - Merging duplicate rows while preserving formulas.
 *      - Moving attended and RSVP entries to their correct locations.
 *      - Importing new events from the Eventbrite sheet to the Contact List.
 *      - Adding new events from the schedule to the Contact List.
 * 3. If the active sheet is not "Contact List":
 *    - Ensures the "Custom Actions" menu is removed.
 *
 * This script enhances usability by providing quick access to key actions
 * for managing event and RSVP data in the Contact List sheet.
 */

  let [sheet, _, _2] = sheetsByName()
  var ui = SpreadsheetApp.getUi();
  if (sheet.getName() === "Contact List") {
    ui.createMenu('Custom Actions')
        .addItem('Sort Attended', 'sortAttendedRows')
        .addItem('Sort RSVPS2+', 'sortRSVPRows')
        .addItem('Merge Duplicates', 'mergeRowsByKeyPreserveAllFormulas')
        .addItem('Move Attended', 'copyAttendedToCorrectLocationAndPreserveRows')
        .addItem('Move RSVP2+', 'copyRSVPToCorrectLocationAndPreserveRows')
        .addItem('Move Rows Eventbrite to Contact List', 'moveRowsFromEventBriteImportToContactList')
        .addItem("New Row and Column", 'importNewEventsFromSchedule')
        .addItem("Move Rows From Facebook Import To Contact List", 'moveRowsFromFaceBookImportToContactList')
        .addItem("Backfill Events Attended & RSVP'd Formulas", 'backfillEventsFormulas')
        .addSeparator()
        .addItem("Generate Data Analysis Graphs", 'createRSVPvsAttendanceChart')
        .addToUi();
    }

  if (sheet.getName() !== "Contact List"){
    ui.createMenu('Custom Actions').removeFromUi();
  }
}
