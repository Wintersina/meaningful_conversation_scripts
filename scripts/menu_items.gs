function onOpen() {
  /**
   * Adds the "Custom Actions" menu to the Google Sheets UI when the
   * spreadsheet is opened with the Contact List active; removes it otherwise.
   */

  let [sheet, _, _2] = sheetsByName()
  var ui = SpreadsheetApp.getUi();
  if (sheet.getName() === "Contact List") {
    ui.createMenu('Custom Actions')
        .addItem('Sort Attended', 'sortAttendedRows')
        .addItem('Sort RSVPS2+', 'sortRSVPRows')
        .addItem('Merge Duplicates', 'mergeRowsByKeyPreserveAllFormulas')
        .addItem('Move Rows Eventbrite to Contact List', 'moveRowsFromEventBriteImportToContactList')
        .addItem("New Row and Column", 'importNewEventsFromSchedule')
        .addItem("Facebook CSV Import…", 'showFacebookCsvImportDialog')
        .addItem("Mark Attendance from Signup Sheet", 'markAttendanceFromSignupSheet')
        .addItem("Email Composer…", 'showEmailComposerDialog')
        .addItem("Bulk Emailer…", 'showBulkEmailerDialog')
        .addItem("No-Email Report…", 'showNoEmailReport')
        //.addItem("Clean Up Signup Origin Columns (K/L/M)", 'cleanupSignupOriginColumns')
        .addSeparator()
        .addItem("Generate Data Analysis Graphs", 'createRSVPvsAttendanceChart')
        .addToUi();
    }

  if (sheet.getName() !== "Contact List"){
    ui.createMenu('Custom Actions').removeFromUi();
  }
}
