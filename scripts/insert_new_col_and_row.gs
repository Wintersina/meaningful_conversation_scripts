function importNewEventsFromSchedule() {
  /**
   * Imports new events from the Schedule sheet into the Contact List sheet.
   *
   * Key behaviors (updated):
   * - Inserts the NEW event column to the RIGHT of column O (so the new column becomes P).
   * - Copies ONLY visual formatting (font/size/colors/borders/alignment/wrap) + width from the previous event column
   *   (which becomes Q after insertion; i.e. the old P).
   * - Does NOT add/replace/modify any conditional formatting rules (no setConditionalFormatRules, no CF copy).
   * - Sets the new date in P6 by taking Q6 + 7 days.
   * - Uses batch writes for header formulas/values to reduce runtime.
   */

  const [contactListSheet, _, scheduleSheet] = sheetsByName();

  // ----------------------------
  // 1) Locate "# Events Attended" column (row 5)
  // ----------------------------
  const lastCol = contactListSheet.getLastColumn();
  const headerRowValues = contactListSheet.getRange(ROW_NUMBERS.ROW_5, 1, 1, lastCol).getValues()[0];

  const attendedIndex0 = headerRowValues.indexOf(COL_CONSTANTS.EVENTS_ATTENDED); // 0-based
  if (attendedIndex0 === -1) return; // FIX: was `== 0` (wrong)

  const attendedCol1 = attendedIndex0 + 1; // 1-based

  // ----------------------------
  // 2) Read schedule rows (skip header row 1) and filter new events
  // ----------------------------
  const scheduleLastRow = scheduleSheet.getLastRow();
  if (scheduleLastRow < 2) return;

  // Columns C & D, starting row 2 (skip header)
  const scheduleData = scheduleSheet.getRange(2, 3, scheduleLastRow - 1, 2).getValues();

  // Existing titles/dates across the contact list header rows
  const contactListTitles = contactListSheet.getRange(ROW_NUMBERS.ROW_7, 1, 1, lastCol).getValues()[0];
  const contactListDates  = contactListSheet.getRange(ROW_NUMBERS.ROW_6, 1, 1, lastCol).getValues()[0];

  const localNormalizeString = (v) => {
    if (v == null) return "";
    return normalizeString(String(v)).toUpperCase();
  };

  const parseDate = (value) => {
    if (!value) return null;

    // Contact list is usually a Date
    if (value instanceof Date) {
      return new Date(value.getFullYear(), value.getMonth(), value.getDate());
    }

    // Schedule is often a string like "Mon, 01/12/26 6:00 PM"
    if (typeof value === "string") {
      const parsed = new Date(value);
      if (!isNaN(parsed)) {
        return new Date(parsed.getFullYear(), parsed.getMonth(), parsed.getDate());
      }
    }

    return null;
  };

  const isWithinWeek = (a, b) => {
    const d1 = parseDate(a);
    const d2 = parseDate(b);
    if (!d1 || !d2) return false;

    const diffDays = Math.abs((d1 - d2) / (1000 * 60 * 60 * 24));
    return diffDays <= 7; // FIX: removed meaningless >= -7 check
  };

  // Find new events (still only inserts the first new one, like your current behavior)
  const newEvents = scheduleData.filter(([date, title]) => {
    if (!title) return false;
    const t = localNormalizeString(title);
    if (!t || SCHEDULE_SHEET_CONSTANTS.SKIP_WORDS.includes(t)) return false;

    const matchFound = contactListTitles.some((existingTitle, idx) => {
      return t === localNormalizeString(existingTitle) && isWithinWeek(date, contactListDates[idx]);
    });

    return !matchFound;
  });

  if (newEvents.length === 0) {
    Logger.log("No new unique events found. Skipping.");
    return;
  }

  const [scheduleDateValue, scheduleTitleValue] = newEvents[0];

  // ----------------------------
  // 3) Insert NEW column to the RIGHT of column O (O=15) => new column becomes P (16)
  //    Copy ONLY visual formatting from the previous P (which becomes Q after insertion) into new P.
  // ----------------------------
  const colOIndex = 15; // O
  contactListSheet.insertColumnAfter(colOIndex);

  const newColIndex = colOIndex + 1;      // P (new column)
  const templateColIndex = newColIndex + 1; // Q (this is the old P after insertion)

  const lastRow = contactListSheet.getLastRow();

  // Copy width from template (old P now Q)
  contactListSheet.setColumnWidth(newColIndex, contactListSheet.getColumnWidth(templateColIndex));

  // Copy ONLY visual formatting (no conditional format rules, no values/formulas)
  // NOTE: PASTE_FORMAT does NOT modify conditional formatting rules; it only pastes static formatting.
  const templateRange = contactListSheet.getRange(1, templateColIndex, lastRow, 1);
  const newColRange   = contactListSheet.getRange(1, newColIndex,      lastRow, 1);
  templateRange.copyTo(newColRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);

  // ----------------------------
  // 4) Set P6 = Q6 + 7 days  (Date row)
  // ----------------------------
  const q6 = contactListSheet.getRange(ROW_NUMBERS.ROW_6, templateColIndex).getValue();
  if (q6 instanceof Date) {
    const newDate = new Date(q6);
    newDate.setDate(newDate.getDate() + 7);
    contactListSheet.getRange(ROW_NUMBERS.ROW_6, newColIndex).setValue(newDate);
  }

  // ----------------------------
  // 5) Set header values + formulas in a batched way (faster)
  //    Contact List semantics:
  //    Row 6 = date (we set above from Q6+7)
  //    Row 7 = title
  //    Row 8 = event id (VLOOKUP)
  //    Row 9 = TTL RSVP
  //    Row 10 = TTL ATTND
  // ----------------------------
  const colLetter = columnToLetter(newColIndex);

  // Row 7 title value
  contactListSheet.getRange(ROW_NUMBERS.ROW_7, newColIndex).setValue(scheduleTitleValue);

  // Row 8/9/10 formulas in one call
  contactListSheet.getRange(ROW_NUMBERS.ROW_8, newColIndex, 3, 1).setFormulas([
    [FORMULAS.V_LOOKUP(colLetter)],
    [FORMULAS.TTL_RSVP(colLetter)],
    [FORMULAS.TTL_ATTND(colLetter)],
  ]);

  // Keep the small font size on TTL rows (optional but matches your prior behavior)
  contactListSheet.getRange(ROW_NUMBERS.ROW_9,  newColIndex).setFontSize(9);
  contactListSheet.getRange(ROW_NUMBERS.ROW_10, newColIndex).setFontSize(9);

  Logger.log("Added new event: " + scheduleTitleValue);

  // ----------------------------
  // 6) Insert rows below the last "Total RSVP'd" block and populate summary formulas
  //    (Speed fixes: move expensive operations OUTSIDE loops; batch formula setting)
  // ----------------------------
  const colBValues = contactListSheet.getRange(1, 2, contactListSheet.getLastRow(), 1).getValues().flat();
  const totalRSVPRowIndex = colBValues.lastIndexOf(COL_CONSTANTS.TOTAL_RSVPD) + 1; // 1-based

  if (totalRSVPRowIndex <= 0) return;

  const insertRowIndex = totalRSVPRowIndex + 1;

  // Insert two rows (these inserts may cause Sheets to auto-adjust CF ranges—this is unavoidable,
  // but we do NOT programmatically alter CF rules anywhere in this script.)
  contactListSheet.insertRows(insertRowIndex, 1);
  contactListSheet.insertRows(insertRowIndex + 5, 1);

  // Match row heights
  const baseHeight = contactListSheet.getRowHeight(totalRSVPRowIndex);
  contactListSheet.setRowHeight(insertRowIndex, baseHeight);
  contactListSheet.setRowHeight(insertRowIndex + 5, baseHeight);

  // Gray styling for the two inserted rows (static formatting only)
  const grayRowRange1 = contactListSheet.getRange(insertRowIndex,     1, 1, contactListSheet.getLastColumn());
  const grayRowRange2 = contactListSheet.getRange(insertRowIndex + 5, 1, 1, contactListSheet.getLastColumn());

  grayRowRange1.setBackground(UI_CONSTANTS.GRAY_BACKGROUND).setWrapStrategy(UI_CONSTANTS.WRAP_STRATEGY).setHorizontalAlignment("left");
  grayRowRange2.setBackground(UI_CONSTANTS.GRAY_BACKGROUND).setWrapStrategy(UI_CONSTANTS.WRAP_STRATEGY).setHorizontalAlignment("left");

  // Labels in column B
  contactListSheet.getRange(insertRowIndex,     2).setValue(scheduleTitleValue);
  contactListSheet.getRange(insertRowIndex + 5, 2).setValue(COL_CONSTANTS.TOTAL_RSVPD);

  // Summary COUNTIF row formulas across event columns
  // Start at O (15), end at "# Events Attended" column
  const startColIndex = 15;           // O
  const endColIndex = attendedCol1;   // inclusive

  const formulaRow = insertRowIndex + 5;
  const numCols = endColIndex - startColIndex + 1;

  // Build formulas for all 3 total rows: Total RSVP'd, Total Attended, Total Attended w/o RSVP
  const rsvpFormulas = [];
  const attendedFormulas = [];
  const attendedNoRsvpFormulas = [];

  for (let col = startColIndex; col <= endColIndex; col++) {
    const letter = columnToLetter(col);
    const rangeStart = insertRowIndex - 30;
    const rangeEnd = insertRowIndex + 4;
    rsvpFormulas.push(FORMULAS.TOTAL_RSVPD_FORMULA(letter, rangeStart, rangeEnd));
    attendedFormulas.push(FORMULAS.TOTAL_ATTENDED_FORMULA(letter, rangeStart, rangeEnd));
    attendedNoRsvpFormulas.push(FORMULAS.TOTAL_ATTENDED_NO_RSVP_FORMULA(letter, rangeStart, rangeEnd));
  }

  // Find the Total Attended and Total Attended w/o RSVP rows (they follow Total RSVP'd)
  // Re-read colB after row insertions since indices shifted
  const colBValuesUpdated = contactListSheet.getRange(1, 2, contactListSheet.getLastRow(), 1).getValues().flat();
  const totalAttendedRowIndex = colBValuesUpdated.lastIndexOf(COL_CONSTANTS.TOTAL_ATTENDED) + 1;
  const totalAttendedNoRsvpRowIndex = colBValuesUpdated.lastIndexOf(COL_CONSTANTS.TOTAL_ATTENDED_NO_RSVP) + 1;

  // Set Total RSVP'd formulas
  contactListSheet.getRange(formulaRow, startColIndex, 1, numCols).setFormulas([rsvpFormulas]);

  // Set Total Attended formulas (if row exists)
  if (totalAttendedRowIndex > 0) {
    contactListSheet.getRange(totalAttendedRowIndex, startColIndex, 1, numCols).setFormulas([attendedFormulas]);
    contactListSheet.getRange(totalAttendedRowIndex, startColIndex, 1, numCols)
      .setNumberFormat(UI_CONSTANTS.NUMBER_FORMAT)
      .setHorizontalAlignment(UI_CONSTANTS.ALIGNMENT_CENTER);
  }

  // Set Total Attended w/o RSVP formulas (if row exists)
  if (totalAttendedNoRsvpRowIndex > 0) {
    contactListSheet.getRange(totalAttendedNoRsvpRowIndex, startColIndex, 1, numCols).setFormulas([attendedNoRsvpFormulas]);
    contactListSheet.getRange(totalAttendedNoRsvpRowIndex, startColIndex, 1, numCols)
      .setNumberFormat(UI_CONSTANTS.NUMBER_FORMAT)
      .setHorizontalAlignment(UI_CONSTANTS.ALIGNMENT_CENTER);
  }

  // Number formatting + center alignment for Total RSVP'd row
  const formulaRange = contactListSheet.getRange(formulaRow, startColIndex, 1, numCols);
  formulaRange.setNumberFormat(UI_CONSTANTS.NUMBER_FORMAT).setHorizontalAlignment(UI_CONSTANTS.ALIGNMENT_CENTER);

  // Bottom border under the last total row (exclude last column like before)
  const bottomBorderRow = totalAttendedNoRsvpRowIndex > 0 ? totalAttendedNoRsvpRowIndex : formulaRow;
  const totalRSVPRowRange = contactListSheet.getRange(bottomBorderRow, 2, 1, contactListSheet.getLastColumn() - 1);
  totalRSVPRowRange.setBorder(false, false, true, false, false, false, UI_CONSTANTS.BORDER_COLOR, UI_CONSTANTS.BORDER_STYLE);

  // ----------------------------
  // 7) Collapsible group: ONE group for rows [insertRowIndex .. insertRowIndex+5]
  //    IMPORTANT: NOT inside any loop.
  // ----------------------------
  const groupStartRow = insertRowIndex + 1 ;
  const groupEndRow = insertRowIndex + 4;
  const groupNumRows = groupEndRow - groupStartRow + 1;

  // Defensive: flatten that range a bit first so reruns don't create nested groups
  contactListSheet.getRange(groupStartRow, 1, groupNumRows, 1).shiftRowGroupDepth(-5);
  contactListSheet.getRange(groupStartRow, 1, groupNumRows, 1).shiftRowGroupDepth(1);
}

