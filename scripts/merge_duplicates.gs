function mergeRowsByKeyPreserveAllFormulas() {

/**
 OLD:formatting =AND(LEN(TRIM($A12))>0, SUMPRODUCT((TRIM($A$1:$A$10207)=TRIM($A12))*1)>1)
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

  // mergedData[key] is an ARRAY of candidate primary rows sharing the same A-key.
  // A duplicate merges into the first candidate whose F and G are compatible
  // (equal, or either side empty — matching the conditional-formatting formula).
  // This means two rows with the same A but conflicting non-empty F or G become
  // two separate primaries instead of falsely merging.
  var mergedData   = {}; // key → [ { dataIndex, fVal, gVal, dirty }, ... ]
  var rowsToDelete = []; // 1-based sheet row numbers of duplicates

  // ── Pass 1: merge all duplicates in memory ────────────────────────────────
  for (var i = 1; i < data.length; i++) { // skip header at index 0
    var key = normalizeByStrippingWhiteSpaceAtTheEnd(data[i][0]);
    if (!key) continue;

    var fVal = data[i][5]; // column F
    var gVal = data[i][6]; // column G

    var candidates = mergedData[key];
    var matched    = null;
    if (candidates) {
      for (var c = 0; c < candidates.length; c++) {
        if (fgCompatible_(candidates[c].fVal, fVal, candidates[c].gVal, gVal)) {
          matched = candidates[c];
          break;
        }
      }
    }

    if (!matched) {
      if (!candidates) candidates = mergedData[key] = [];
      candidates.push({ dataIndex: i, fVal: fVal, gVal: gVal, dirty: false });
    } else {
      var fi = matched.dataIndex;
      matched.dirty = true; // this row received at least one merge

      for (var col = 0; col < data[i].length; col++) {
        if (skipColsSet.has(col)) continue;

        var existingValue = data[fi][col];
        var newValue      = data[i][col];

        // Promote a formula from the duplicate if the primary row has none
        if (!formulas[fi][col] && formulas[i][col]) {
          formulas[fi][col] = formulas[i][col];
        }

        // Event columns (O onward): pick by RSVP priority, never concatenate
        if (col >= HELPER_CONSTANTS.EVENT_NAMES_START_COL - 1) {
          var winner = pickByRsvpPriority_(existingValue, newValue);
          if (winner !== undefined) {
            data[fi][col] = winner;
          }
        } else {
          // Non-event columns: original merge logic.
          // For F/G specifically, populated > empty falls into the first branch,
          // and equal-populated values skip the concat branch — so F and G end up
          // holding the non-empty value without concatenation. The concat guard
          // uses loose equality so a case/whitespace-only difference (e.g. the
          // same email with a different capitalization) keeps the primary's
          // value instead of producing "a@x.com, A@x.com".
          if (!existingValue || existingValue === "-") {
            if (newValue && newValue !== "-") {
              data[fi][col] = newValue;
            }
          } else if (!eqLoose_(existingValue, newValue) && newValue && newValue !== "-") {
            data[fi][col] = existingValue + ", " + newValue;
          }
        }
      }

      // Update the candidate's F/G so subsequent rows compare against the
      // now-populated values (empty + populated → populated takes precedence).
      if (isEmptyFgVal_(matched.fVal) && !isEmptyFgVal_(fVal)) matched.fVal = fVal;
      if (isEmptyFgVal_(matched.gVal) && !isEmptyFgVal_(gVal)) matched.gVal = gVal;

      rowsToDelete.push(i + 1); // 1-based sheet row
    }
  }

  // ── Pass 2: write back only dirty rows (rows that had duplicates merged in) ──
  var allEntries = [];
  Object.values(mergedData).forEach(function(candidates) {
    candidates.forEach(function(entry) { allEntries.push(entry); });
  });
  allEntries.forEach(function(entry) {
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

/**
 * Given two RSVP cell values from duplicate rows, returns the higher-priority one.
 * Priority (highest first):
 *   1. attended: yes  (any rsvp prefix)
 *   2. attended: no   (any rsvp prefix)
 *   3. attended: ?    (any rsvp prefix)
 *   4. dash / empty
 * Returns undefined if neither value is meaningful.
 */
function pickByRsvpPriority_(a, b) {
  // Rank: lower number = higher priority
  var RSVP_PRIORITY = [
    RSVP_DROP_DOWN_CONSTANTS.YES_ATTENDED_YES,   // 0
    RSVP_DROP_DOWN_CONSTANTS.MAYBE_ATTENDED_YES,  // 1
    RSVP_DROP_DOWN_CONSTANTS.NO_ATTENDED_YES,     // 2
    RSVP_DROP_DOWN_CONSTANTS.DASH_ATTENDED_YES,   // 3
    RSVP_DROP_DOWN_CONSTANTS.YES_ATTENDED_NO,     // 4
    RSVP_DROP_DOWN_CONSTANTS.NO_ATTENDED_NO,      // 5
    RSVP_DROP_DOWN_CONSTANTS.YES_ATTENDED,        // 6
    RSVP_DROP_DOWN_CONSTANTS.MAYBE_ATTENDED,      // 7
    RSVP_DROP_DOWN_CONSTANTS.NO_ATTENDED,         // 8
    RSVP_DROP_DOWN_CONSTANTS.DASH,                // 9
    RSVP_DROP_DOWN_CONSTANTS.DOUBLE_DASH          // 10
  ];

  var aEmpty = !a || a === "-" || a === "--";
  var bEmpty = !b || b === "-" || b === "--";

  if (aEmpty && bEmpty) return undefined;
  if (aEmpty) return b;
  if (bEmpty) return a;

  var aRank = RSVP_PRIORITY.indexOf(a);
  var bRank = RSVP_PRIORITY.indexOf(b);

  // If a value isn't in the priority list, treat it as lowest
  if (aRank === -1) aRank = 999;
  if (bRank === -1) bRank = 999;

  return aRank <= bRank ? a : b;
}

// Treats null/undefined and the empty string as "empty" for F/G wildcard logic,
// mirroring the conditional-formatting formula's `$F12=""` check.
function isEmptyFgVal_(v) {
  return v == null || v === "";
}

// Case- and whitespace-insensitive equality for the F/G identity columns.
// Mirrors the row-key normalization (normalizeByStrippingWhiteSpaceAtTheEnd),
// so the merge gate and the actual merge agree: the same email with a
// different capitalization (b.s.klump@ vs B.s.klump@) is treated as equal.
function eqLoose_(a, b) {
  return String(a == null ? "" : a).trim().toLowerCase() ===
         String(b == null ? "" : b).trim().toLowerCase();
}

// F and G are compatible for merging when each side is either equal or empty —
// i.e. the formula's ($F1=$F12) + ($F1="") + ($F12="") > 0 evaluated for both
// columns. Uses loose equality (case/whitespace-insensitive) so trivial casing
// differences in F/G don't split one person into two un-merged primaries.
function fgCompatible_(aF, bF, aG, bG) {
  var fOk = isEmptyFgVal_(aF) || isEmptyFgVal_(bF) || eqLoose_(aF, bF);
  var gOk = isEmptyFgVal_(aG) || isEmptyFgVal_(bG) || eqLoose_(aG, bG);
  return fOk && gOk;
}
