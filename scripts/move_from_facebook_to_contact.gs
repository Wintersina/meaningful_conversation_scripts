function moveRowsFromFaceBookImportToContactList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const contactListSheet = ss.getSheetByName(SHEET_NAMES.CONTACT_LIST);
  const facebookSheet = ss.getSheetByName(SHEET_NAMES.FACEBOOK);

  if (!contactListSheet) throw new Error(ERROR_MESSAGES.CONTACT_LIST_NOT_FOUND);
  if (!facebookSheet) throw new Error("'FaceBook Import' sheet not found.");

  // ----------------------------------------
  // Guard: exit if Facebook has no data
  // ----------------------------------------
  const fbLastRow = facebookSheet.getLastRow();
  if (fbLastRow <= 1) return;

  const contactLastRow = contactListSheet.getLastRow();
  const contactLastCol = contactListSheet.getLastColumn();

  // ----------------------------------------
  // Load Facebook data (A..C, rows 2..N)
  // ----------------------------------------
  const fbData = facebookSheet
    .getRange(
      HELPER_CONSTANTS.FIRST_DATA_ROW,
      1,
      fbLastRow - 1,
      HELPER_CONSTANTS.FACEBOOK_COLUMN_COUNT
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

  const headerRow = contactListSheet
    .getRange(ROW_NUMBERS.ROW_5, 1, 1, contactLastCol)
    .getValues()[0]; // Counters row

  const rsvpCol0 = headerRow.indexOf(COL_CONSTANTS.EVENTS_RSVPD);
  const attendedCol0 = headerRow.indexOf(COL_CONSTANTS.EVENTS_ATTENDED);
  if (rsvpCol0 === -1 || attendedCol0 === -1) return;

  const rsvpCol = rsvpCol0 + 1;
  const attendedCol = attendedCol0 + 1;
  const attendedColLetter = columnToLetter(attendedCol);

  // ----------------------------------------
  // Build RSVP dropdown validation once
  // ----------------------------------------
  const rsvpValidationRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(Object.values(RSVP_DROP_DOWN_CONSTANTS), true)
    .setAllowInvalid(false)
    .build();

  // ----------------------------------------
  // Build a normalized lookup for Contact List Column B (event blocks)
  // Bottom-most match wins, like Eventbrite logic.
  // ----------------------------------------
  const normalizedContactTitles = contactColB.map(v =>
    normalizeByStrippingWhiteSpaceAtTheEnd(stripParen_(v))
  );

  // ----------------------------------------
  // Group Facebook rows by cleaned event name (Title col C)
  // ----------------------------------------
  const groupedEvents = {};
  fbData.forEach((row, idx) => {
    const rawTitle = row[2]; // Column C
    if (!rawTitle) return;

    const cleanedTitle = stripParen_(normalizeString(String(rawTitle))).trim();
    const key = normalizeByStrippingWhiteSpaceAtTheEnd(cleanedTitle);
    if (!key) return;

    if (!groupedEvents[key]) groupedEvents[key] = [];
    groupedEvents[key].push({ row, idx, cleanedTitle, rawTitle });
  });

  // ----------------------------------------
  // Build work items (resolved positions)
  // ----------------------------------------
  const workItems = [];

  Object.entries(groupedEvents).forEach(([eventKey, rows]) => {
    const matchIndex = normalizedContactTitles.lastIndexOf(eventKey);
    if (matchIndex === -1) return;

    const matchRow = matchIndex + 1;

    // Prefer LEFTMOST event column (starting at O) for RSVP placement
    const cleanedTitleForHeader = rows[0].cleanedTitle; // stripped version
    const titleCol = findPreferredEventTitleColumn_(
      titlesRow,
      cleanedTitleForHeader,
      15,               // Column O
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
    // Build values we can set:
    // - C First, D Last
    // - K Platform = FB
    // - L Original Signup Event Title (rawTitle)
    // ----------------------------------------
    const firstNames = [];
    const lastNames = [];
    const platforms = [];
    const originalTitles = [];

    // RSVP values depend on Status
    const rsvpValues = [];

    rows.forEach(({ row, rawTitle }) => {
      const fullName = String(row[0] || "").trim(); // Column A
      const status = String(row[1] || "").trim();   // Column B

      const split = splitName_(fullName);
      firstNames.push([split.first]);
      lastNames.push([split.last]);

      platforms.push(["FB"]);
      originalTitles.push([rawTitle || ""]);

      rsvpValues.push([statusToRsvpValue_(status)]);
    });

    // Set First/Last (C, D)
    contactListSheet.getRange(insertAt, 3, rowCount, 1).setValues(firstNames);
    contactListSheet.getRange(insertAt, 4, rowCount, 1).setValues(lastNames);

    // Set Platform (K=11), Title (L=12)
    contactListSheet.getRange(insertAt, 11, rowCount, 1).setValues(platforms);
    contactListSheet.getRange(insertAt, 12, rowCount, 1).setValues(originalTitles);

    // ----------------------------------------
    // Insert formulas (batched)
    // ----------------------------------------
    const nameFormulas = [];
    const attendedFormulas = [];
    const rsvpCountFormulas = [];

    for (let i = 0; i < rowCount; i++) {
      const rowNum = insertAt + i;
      nameFormulas.push([FORMULAS.CONCATENATE_NAME(rowNum)]);
      attendedFormulas.push([FORMULAS.COUNT_ATTENDED(rowNum, attendedColLetter)]);
      rsvpCountFormulas.push([FORMULAS.COUNT_RSVP(rowNum, attendedColLetter)]);
    }

    contactListSheet.getRange(insertAt, 1, rowCount, 1).setFormulas(nameFormulas);
    contactListSheet.getRange(insertAt, attendedCol, rowCount, 1).setFormulas(attendedFormulas);
    contactListSheet.getRange(insertAt, rsvpCol, rowCount, 1).setFormulas(rsvpCountFormulas);

    // ----------------------------------------
    // RSVP dropdown + dash fill (like Eventbrite)
    // ----------------------------------------
    if (titleCol) {
      const rsvpRange = contactListSheet.getRange(insertAt, titleCol, rowCount, 1);
      rsvpRange.setDataValidation(rsvpValidationRule);
      rsvpRange.setValues(rsvpValues);

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

    // Track rows for deletion from Facebook Import
    rows.forEach(r =>
      rowsToDelete.push(r.idx + HELPER_CONSTANTS.FIRST_DATA_ROW)
    );
  });

  // ----------------------------------------
  // Delete Facebook rows bottom → top
  // ----------------------------------------
  rowsToDelete
    .sort((a, b) => b - a)
    .forEach(r => facebookSheet.deleteRow(r));
}

/**
 * Strip trailing/inner parenthetical bits like: "Title (Free Event)" -> "Title"
 */
function stripParen_(value) {
  if (!value) return "";
  return String(value).replace(/\s*\(.*?\)\s*/g, "").trim();
}

/**
 * Split a full name into first + last.
 * - "Mary" => first="Mary", last=""
 * - "Mary Jane Tkacz" => first="Mary", last="Jane Tkacz"
 */
function splitName_(fullName) {
  const cleaned = String(fullName || "").trim().replace(/\s+/g, " ");
  if (!cleaned) return { first: "", last: "" };

  const parts = cleaned.split(" ");
  if (parts.length === 1) return { first: parts[0], last: "" };

  return { first: parts[0], last: parts.slice(1).join(" ") };
}

/**
 * Map Facebook status -> your RSVP dropdown value.
 */
function statusToRsvpValue_(status) {
  const s = normalizeByStrippingWhiteSpaceAtTheEnd(status) || "";

  if (s === "maybe" || s === "interested") {
    return RSVP_DROP_DOWN_CONSTANTS.MAYBE_ATTENDED;
  }

  if (s === "going" || s === "yes") {
    return RSVP_DROP_DOWN_CONSTANTS.YES_ATTENDED;
  }

  if (s === "no" || s === "declined" || s === "not going") {
    return RSVP_DROP_DOWN_CONSTANTS.NO_ATTENDED;
  }

  // safest default:
  return RSVP_DROP_DOWN_CONSTANTS.MAYBE_ATTENDED;
}
