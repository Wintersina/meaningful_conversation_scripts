/**
 * Backfills fresh "# Events Attended" and "# Events RSVP'd" formulas for every
 * contact row in the Contact List sheet.
 *
 * A "contact row" is identified by having the CONCATENATE_NAME formula in Column A
 * (e.g. =IF(ISTEXT(C{n}),CONCATENATE(C{n}," ",D{n}),"")).
 *
 * This is needed when new event columns are added and the column letters in the
 * COUNTIF range (O{n}:{letter}{n}) fall behind. Running this script re-writes
 * both formulas with the current correct column letter derived from the header row.
 *
 * Rows are processed in consecutive groups for batch efficiency.
 */
function backfillEventsFormulas() {
  Logger.log("starting backfillEventsFormulas");

  const [sheet] = sheetsByName();

  const lastRow  = sheet.getLastRow();
  const lastCol  = sheet.getLastColumn();

  // ----------------------------------------
  // Find attended and rsvp column indices
  // ----------------------------------------
  const headerRow     = sheet.getRange(ROW_NUMBERS.ROW_5, 1, 1, lastCol).getValues()[0];
  const attendedCol0  = headerRow.indexOf(COL_CONSTANTS.EVENTS_ATTENDED);
  const rsvpCol0      = headerRow.indexOf(COL_CONSTANTS.EVENTS_RSVPD);

  if (attendedCol0 === -1 || rsvpCol0 === -1) {
    Logger.log("backfillEventsFormulas: could not find attended/rsvp columns — aborting");
    return;
  }

  const attendedCol       = attendedCol0 + 1; // 1-based
  const rsvpCol           = rsvpCol0 + 1;     // 1-based
  const attendedColLetter = columnToLetter(attendedCol);

  // ----------------------------------------
  // Read ALL column-A formulas in one call
  // ----------------------------------------
  const colAFormulas = sheet.getRange(1, 1, lastRow, 1).getFormulas();

  // ----------------------------------------
  // Group consecutive rows that have the name formula into runs so we can
  // use setFormulas() on each contiguous block (far fewer API calls)
  // ----------------------------------------
  const groups = [];
  let currentGroup = null;

  for (let i = 0; i < colAFormulas.length; i++) {
    const formula  = colAFormulas[i][0];
    const hasName  = formula && formula.toUpperCase().includes("CONCATENATE");
    const rowNum   = i + 1; // 1-based

    if (hasName) {
      if (currentGroup && currentGroup.start + currentGroup.rows.length === rowNum) {
        // Consecutive — extend the current group
        currentGroup.rows.push(rowNum);
      } else {
        // Start a new group
        if (currentGroup) groups.push(currentGroup);
        currentGroup = { start: rowNum, rows: [rowNum] };
      }
    } else {
      if (currentGroup) {
        groups.push(currentGroup);
        currentGroup = null;
      }
    }
  }
  if (currentGroup) groups.push(currentGroup);

  // ----------------------------------------
  // Batch-write formulas per group
  // ----------------------------------------
  let totalUpdated = 0;

  groups.forEach(function(group) {
    const numRows        = group.rows.length;
    const attendedFmls  = [];
    const rsvpFmls      = [];

    group.rows.forEach(function(rowNum) {
      attendedFmls.push([FORMULAS.COUNT_ATTENDED(rowNum, attendedColLetter)]);
      rsvpFmls.push([FORMULAS.COUNT_RSVP(rowNum, attendedColLetter)]);
    });

    sheet.getRange(group.start, attendedCol, numRows, 1).setFormulas(attendedFmls);
    sheet.getRange(group.start, rsvpCol,     numRows, 1).setFormulas(rsvpFmls);

    totalUpdated += numRows;
  });

  Logger.log("backfillEventsFormulas: updated " + totalUpdated + " rows across " + groups.length + " groups");
}
