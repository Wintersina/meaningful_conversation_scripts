/**
 * Moves rows from the Eventbrite Import sheet into the Contact List sheet.
 *
 * PERFORMANCE-OPTIMIZED VERSION
 *
 * Key guarantees:
 * - Bottom-to-top processing to avoid row index corruption.
 * - Batch row insertion per event (fast).
 * - No conditional formatting rules are added, removed, or modified.
 * - Collapsed row groups are expanded once per event before insertion.
 * - RSVP placement prioritizes the LEFTMOST event column (closest to column O).
 * - Attendee rows are inserted TWO rows below the event title.
 */
function moveRowsFromEventBriteImportToContactList() {
  const [contactListSheet, eventbriteSheet] = sheetsByName();

  // ----------------------------------------
  // Guard: exit if Eventbrite has no data
  // ----------------------------------------
  const ebLastRow = eventbriteSheet.getLastRow();
  if (ebLastRow <= 1) return;

  const contactLastRow = contactListSheet.getLastRow();
  const contactLastCol = contactListSheet.getLastColumn();

  // ----------------------------------------
  // Load Eventbrite data (A..L, rows 2..N)
  // ----------------------------------------
  const eventbriteData = eventbriteSheet
    .getRange(
      HELPER_CONSTANTS.FIRST_DATA_ROW,
      1,
      ebLastRow - 1,
      HELPER_CONSTANTS.EVENTBRITE_COLUMN_COUNT
    )
    .getValues();

  // ----------------------------------------
  // Preload Contact List metadata
  // ----------------------------------------
  const contactColB = contactListSheet
    .getRange(1, 2, contactLastRow, 1)
    .getValues()
    .flat(); // Event block titles (Column B)

  const titlesRow = contactListSheet
    .getRange(ROW_NUMBERS.ROW_7, 1, 1, contactLastCol)
    .getValues()[0]; // Event headers across columns

  const rsvpCol     = findColMarker_(contactListSheet, MARKER_KEYS.EVENTS_RSVPD,    COL_CONSTANTS.EVENTS_RSVPD);    // 1-based
  const attendedCol = findColMarker_(contactListSheet, MARKER_KEYS.EVENTS_ATTENDED, COL_CONSTANTS.EVENTS_ATTENDED); // 1-based
  if (rsvpCol === -1 || attendedCol === -1) return;
  const attendedColLetter = columnToLetter(attendedCol);

  // ----------------------------------------
  // Build RSVP dropdown validation once
  // ----------------------------------------
  const rsvpValidationRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(Object.values(RSVP_DROP_DOWN_CONSTANTS), true)
    .setAllowInvalid(false)
    .build();

  // ----------------------------------------
  // Group Eventbrite rows by cleaned event name
  // ----------------------------------------
  const groupedEvents = {};
  eventbriteData.forEach((row, idx) => {
    const rawEvent = row[10]; // Column K
    if (!rawEvent) return;

    const cleaned = normalizeString(String(rawEvent))
      .replace(/\s*\(.*?\)\s*/g, "")
      .trim();

    if (!groupedEvents[cleaned]) groupedEvents[cleaned] = [];
    groupedEvents[cleaned].push({ row, idx });
  });

  // ----------------------------------------
  // Build work items (resolved positions)
  // ----------------------------------------
  const workItems = [];

  Object.entries(groupedEvents).forEach(([eventName, rows]) => {
    // Bottom-most matching event block wins
    const matchIndex = contactColB.lastIndexOf(eventName);
    if (matchIndex === -1) return;

    const matchRow = matchIndex + 1;

    // Choose LEFTMOST matching event column starting at O
    const titleCol = findPreferredEventTitleColumn_(
      titlesRow,
      eventName,
      15,                  // Column O
      attendedCol - 1
    );

    workItems.push({ matchRow, titleCol, rows });
  });

  // ----------------------------------------
  // Sort bottom → top to preserve row indexes
  // ----------------------------------------
  workItems.sort((a, b) => b.matchRow - a.matchRow);

  const rowsToDelete = [];

  // ----------------------------------------
  // Execute inserts (batched per event)
  // ----------------------------------------
  workItems.forEach(({ matchRow, titleCol, rows }) => {
    const insertAt = matchRow + 2; // ✅ INSERT TWO ROWS BELOW TITLE
    const rowCount = rows.length;

    // Expand collapsed groups once per event
    expandRowGroupsAtRow_(contactListSheet, insertAt);
    contactListSheet.showRows(insertAt);

    // Insert all attendee rows in one operation
    contactListSheet.insertRowsBefore(insertAt, rowCount);

    // ----------------------------------------
    // Insert attendee values (Columns C..)
    // ----------------------------------------
    const values = rows.map(r =>
      r.row.slice(1, HELPER_CONSTANTS.EVENTBRITE_COLUMN_COUNT)
    );

    contactListSheet
      .getRange(insertAt, 3, rowCount, values[0].length)
      .setValues(values);

    // ----------------------------------------
    // Insert formulas (batched)
    // ----------------------------------------
    const nameFormulas = [];
    const attendedFormulas = [];
    const rsvpFormulas = [];

    for (let i = 0; i < rowCount; i++) {
      const rowNum = insertAt + i;
      nameFormulas.push([FORMULAS.CONCATENATE_NAME(rowNum)]);
      attendedFormulas.push([FORMULAS.COUNT_ATTENDED(rowNum, attendedColLetter)]);
      rsvpFormulas.push([FORMULAS.COUNT_RSVP(rowNum, attendedColLetter)]);
    }

    contactListSheet.getRange(insertAt, 1, rowCount, 1).setFormulas(nameFormulas);
    contactListSheet.getRange(insertAt, attendedCol, rowCount, 1).setFormulas(attendedFormulas);
    contactListSheet.getRange(insertAt, rsvpCol, rowCount, 1).setFormulas(rsvpFormulas);

    // ----------------------------------------
    // RSVP dropdown + dash fill
    // ----------------------------------------
    if (titleCol) {
      const rsvpRange = contactListSheet.getRange(insertAt, titleCol, rowCount, 1);
      rsvpRange.setDataValidation(rsvpValidationRule);
      rsvpRange.setValues(
        Array(rowCount).fill([RSVP_DROP_DOWN_CONSTANTS.YES_ATTENDED])
      );

      const dashStart = titleCol + 1;
      const dashEnd = attendedCol - 1;
      if (dashStart <= dashEnd) {
        const width = dashEnd - dashStart + 1;
        contactListSheet
          .getRange(insertAt, dashStart, rowCount, width)
          .setValues(Array(rowCount).fill(Array(width).fill("-")));
      }
    }

    // ----------------------------------------
    // Static formatting (one call)
    // ----------------------------------------
    contactListSheet
      .getRange(insertAt, 1, rowCount, contactLastCol)
      .setFontSize(UI_CONSTANTS.FONT_SIZE)
      .setFontFamily(UI_CONSTANTS.FONT_STYLE)
      .setHorizontalAlignment(UI_CONSTANTS.ALIGNMENT_CENTER);

    // Track rows for deletion from Eventbrite
    rows.forEach(r =>
      rowsToDelete.push(r.idx + HELPER_CONSTANTS.FIRST_DATA_ROW)
    );
  });

  // ----------------------------------------
  // Delete Eventbrite rows bottom → top
  // ----------------------------------------
  rowsToDelete
    .sort((a, b) => b - a)
    .forEach(r => eventbriteSheet.deleteRow(r));
}

/**
 * Expands any collapsed row groups at a given row index.
 * This prevents inserts from landing inside hidden groups.
 */
function expandRowGroupsAtRow_(sheet, row) {
  try {
    const depth = sheet.getRowGroupDepth(row);
    for (let d = depth; d >= 1; d--) {
      const group = sheet.getRowGroup(row, d);
      if (group) group.expand();
    }
  } catch (e) {
    // Fail silently if grouping APIs are unavailable
  }
}

/**
 * Returns the LEFTMOST event column (closest to column O)
 * whose header text matches the given event title.
 */
function findPreferredEventTitleColumn_(titlesRow, title, firstCol, lastCol) {
  const target = String(title || "").trim();
  if (!target) return null;

  for (let c = firstCol - 1; c <= lastCol - 1; c++) {
    if (titlesRow[c] && String(titlesRow[c]).trim() === target) {
      return c + 1; // 1-based
    }
  }
  return null;
}
