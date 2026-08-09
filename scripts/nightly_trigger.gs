function nightlyTrigger() {
  Logger.log("starting nightlyTrigger");

  // Add any new Schedule events first so the EventBrite move below can find
  // their event blocks/columns. No-ops when the Schedule has nothing new.
  importNewEventsFromSchedule()

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
