/**
 * Builds the spreadsheet's custom menus on open:
 *
 *   Custom Actions — sheet maintenance (sorting, merging, imports, charts)
 *   Email          — everything that sends or reports on email: the Email
 *                    Composer, the Bulk Emailer, and the No-Email Report
 *
 * Email lives in its own top-level menu rather than buried in Custom Actions,
 * because that's the work done most often and by the most people.
 */
function onOpen() {
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
        //.addItem("Clean Up Signup Origin Columns (K/L/M)", 'cleanupSignupOriginColumns')
        .addSeparator()
        .addItem("Generate Data Analysis Graphs", 'createRSVPvsAttendanceChart')
        .addToUi();

    ui.createMenu('Email')
        .addItem("Email Composer…", 'showEmailComposerDialog')
        .addItem("Bulk Emailer…", 'showBulkEmailerDialog')
        .addSeparator()
        .addItem("No-Email Report…", 'showNoEmailReport')
        .addToUi();
    }

  if (sheet.getName() !== "Contact List"){
    ui.createMenu('Custom Actions').removeFromUi();
    ui.createMenu('Email').removeFromUi();
  }
}
