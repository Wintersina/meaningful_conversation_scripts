function nightlyTrigger() {
  Logger.log("starting nightlyTrigger");

  moveRowsFromEventBriteImportToContactList()
  mergeRowsByKeyPreserveAllFormulas()
  copyAttendedToCorrectLocationAndPreserveRows()
  copyRSVPToCorrectLocationAndPreserveRows()
  sortAttendedRows()
  sortRSVPRows()

  syncEventsToGoogleCalendar()

  Logger.log("ending nightlyTrigger");

}
