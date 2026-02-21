function nightlyTrigger() {
  Logger.log("starting nightlyTrigger");

  moveRowsFromEventBriteImportToContactList()
  mergeRowsByKeyPreserveAllFormulas()
  copyAttendedToCorrectLocationAndPreserveRows()
  copyRSVPToCorrectLocationAndPreserveRows()
  sortAttendedRows()
  sortRSVPRows()

  Logger.log("ending nightlyTrigger");

}
