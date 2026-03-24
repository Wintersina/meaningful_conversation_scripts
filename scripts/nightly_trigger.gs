function nightlyTrigger() {
  Logger.log("starting nightlyTrigger");

  moveRowsFromEventBriteImportToContactList()
  mergeRowsByKeyPreserveAllFormulas()
  copyAttendedToCorrectLocationAndPreserveRows()
  copyRSVPToCorrectLocationAndPreserveRows()
  sortAttendedRows()
  sortRSVPRows()

  markNoShows()
  syncEventsToGoogleCalendar()

  Logger.log("ending nightlyTrigger");

}
