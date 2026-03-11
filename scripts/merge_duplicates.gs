function mergeRowsByKeyPreserveAllFormulas() {

/**
 * Merges duplicate rows in the sheet while preserving all formulas.
 *
 * This function:
 * 1. Retrieves all data and formulas from the sheet (2 API calls total for reads).
 * 2. Identifies duplicate rows based on the key in Column A.
 * 3. Merges values from duplicate rows into a single row in memory while:
 *    - Preserving formulas.
 *    - Concatenating unique values.
 *    - Skipping specific columns that should not be merged.
 * 4. Refreshes the attended/rsvp formula column letters so they never go stale.
 * 5. Writes back ONLY the rows that actually had duplicates merged into them.
 * 6. Deletes duplicate rows, grouping consecutive runs into single deleteRows() calls.
 *
 * Key Implementation Details:
 * - `skipColsSet`: Set (O(1) lookup) of columns that should not be merged.
 * - Row 5 header is read from the already-loaded `data[]` — no extra API call.
 * - Only rows that received a merge are written back, avoiding touching the full range.
 */
  Logger.log("starting mergeRowsByKeyPreserveAllFormulas");
  let [sheet] = sheetsByName();

  var range    = sheet.getDataRange();
  var data     = range.getValues();   // 1 API call — all cell values
  var formulas = range.getFormulas(); // 1 API call — all cell formulas

  // ── Find column indices from the already-loaded data (no extra API call) ──
  const headerArr        = data[ROW_NUMBERS.ROW_5 - 1]; // 0-based into data[]
  const RSVPColIndex     = headerArr.indexOf(COL_CONSTANTS.EVENTS_RSVPD);
  const AttendedColIndex = headerArr.indexOf(COL_CONSTANTS.EVENTS_ATTENDED);

  if (RSVPColIndex === -1 || AttendedColIndex === -1) {
    Logger.log("mergeRowsByKeyPreserveAllFormulas: could not find attended/rsvp columns — aborting");
    return;
  }

  const attendedColLetter = columnToLetter(AttendedColIndex + 1);

  // Set gives O(1) lookup vs array's O(n) — matters inside the inner loop
  const skipColsSet = new Set([0, 1, 2, 3, 9, 11, RSVPColIndex, AttendedColIndex, RSVPColIndex + 1]);

  var mergedData   = {}; // key → { dataIndex, dirty }
  var rowsToDelete = []; // 1-based sheet row numbers of duplicates

  // ── Pass 1: merge all duplicates in memory ────────────────────────────────
  for (var i = 1; i < data.length; i++) { // skip header at index 0
    var key = normalizeByStrippingWhiteSpaceAtTheEnd(data[i][0]);
    if (!key) continue;

    if (!mergedData[key]) {
      mergedData[key] = { dataIndex: i, dirty: false };
    } else {
      var fi = mergedData[key].dataIndex;
      mergedData[key].dirty = true; // this row received at least one merge

      for (var col = 0; col < data[i].length; col++) {
        if (skipColsSet.has(col)) continue;

        var existingValue = data[fi][col];
        var newValue      = data[i][col];

        // Promote a formula from the duplicate if the primary row has none
        if (!formulas[fi][col] && formulas[i][col]) {
          formulas[fi][col] = formulas[i][col];
        }

        // Merge values
        if (!existingValue || existingValue === "-") {
          if (newValue && newValue !== "-") {
            data[fi][col] = newValue;
          }
        } else if (existingValue !== newValue && newValue && newValue !== "-") {
          data[fi][col] = existingValue + ", " + newValue;
        }
      }

      rowsToDelete.push(i + 1); // 1-based sheet row
    }
  }

  // ── Pass 2: write back only dirty rows (rows that had duplicates merged in) ──
  Object.values(mergedData).forEach(function(entry) {
    if (!entry.dirty) return; // skip rows that had no duplicates

    var i      = entry.dataIndex;
    var rowNum = i + 1; // 1-based sheet row

    // Refresh attended/rsvp formula column letters before writing
    if (formulas[i][AttendedColIndex]) {
      formulas[i][AttendedColIndex] = FORMULAS.COUNT_ATTENDED(rowNum, attendedColLetter);
    }
    if (formulas[i][RSVPColIndex]) {
      formulas[i][RSVPColIndex] = FORMULAS.COUNT_RSVP(rowNum, attendedColLetter);
    }

    // Write merged values for this row
    sheet.getRange(rowNum, 1, 1, data[i].length).setValues([data[i]]);

    // Write back formulas for this row (only cells that have formulas)
    formulas[i].forEach(function(formula, colIndex) {
      if (formula) {
        sheet.getRange(rowNum, colIndex + 1).setFormula(formula);
      }
    });
  });

  // ── Delete duplicate rows, batching consecutive runs ──────────────────────
  // Sorting descending means deletions don't shift the indices of later rows.
  rowsToDelete.sort(function(a, b) { return b - a; });

  var idx = 0;
  while (idx < rowsToDelete.length) {
    var hi = rowsToDelete[idx]; // top of a consecutive block
    var lo = hi;
    var j  = idx + 1;
    while (j < rowsToDelete.length && rowsToDelete[j] === lo - 1) {
      lo = rowsToDelete[j]; // extend the block downward
      j++;
    }
    sheet.deleteRows(lo, hi - lo + 1); // one call per consecutive block
    idx = j;
  }

  Logger.log("ending mergeRowsByKeyPreserveAllFormulas");
}
