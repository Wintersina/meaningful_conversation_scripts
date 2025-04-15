function nightlyTrigger() {
  Logger.log("starting nightlyTrigger");

  mergeRowsByKeyPreserveAllFormulas()
  copyAttendedToCorrectLocationAndPreserveRows()
  copyRSVPToCorrectLocationAndPreserveRows()
  sortAttendedRows()
  sortRSVPRows()

  Logger.log("ending nightlyTrigger");

}
