function onOpen() {
  // Create a custom menu to run the scripts
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
        .addToUi();
    }

  if (sheet.getName() !== "Contact List"){
    ui.createMenu('Custom Actions').removeFromUi();
  }
}
