/**
 * One-shot cleanup for the Original Signup columns K/L/M (Platform, Event
 * Title, Event Code) in the Contact List. Earlier merges concatenated these
 * with ", " on every duplicate merge ("EB, FB, Walk-in"); the merge no longer
 * does that (J/K/L/M are in the merge skip-set), and this script repairs the
 * rows that already got conjoined by keeping only the left-most (original)
 * value.
 *
 * Column L needs care: event titles can legitimately contain commas
 * ("One God, Many Paths"), so a naive split would truncate real titles.
 * L is therefore matched against the known event titles in Row 7 — the
 * longest comma-joined prefix that equals a known title wins; only when
 * nothing matches does the plain first comma-part get kept.
 *
 * Run from Custom Actions → "Clean Up Signup Origin Columns (K/L/M)".
 */
function cleanupSignupOriginColumns() {
  Logger.log("starting cleanupSignupOriginColumns");
  var [contactListSheet] = sheetsByName();

  var lastRow = contactListSheet.getLastRow();
  var dataStartRow = ROW_NUMBERS.ROW_12;
  var numRows = lastRow - dataStartRow + 1;
  if (numRows <= 0) {
    Logger.log("No data rows to clean.");
    return;
  }

  // Known event titles from Row 7 (event columns start at O), normalized
  var lastCol = contactListSheet.getLastColumn();
  var titlesRow = contactListSheet.getRange(ROW_NUMBERS.ROW_7, 1, 1, lastCol).getValues()[0];
  var knownTitles = new Set();
  for (var c = HELPER_CONSTANTS.EVENT_NAMES_START_COL - 1; c < titlesRow.length; c++) {
    var t = normalizeByStrippingWhiteSpaceAtTheEnd(titlesRow[c]);
    if (t) knownTitles.add(t);
  }

  // K/L/M = 1-based columns 11..13, one bulk read
  var firstCol = COLUMN_INDEX.SIGNUP_PLATFORM + 1; // K = 11
  var range = contactListSheet.getRange(dataStartRow, firstCol, numRows, 3);
  var values = range.getValues();

  var changed = 0;
  for (var r = 0; r < values.length; r++) {
    var platform = cleanLeftmost_(values[r][0]);
    var title = cleanTitleKeepingKnown_(values[r][1], knownTitles);
    var code = cleanLeftmost_(values[r][2]);

    if (platform !== values[r][0] || title !== values[r][1] || code !== values[r][2]) {
      Logger.log("Row " + (dataStartRow + r) + ": [" + values[r].join(" | ") + "] → [" +
        platform + " | " + title + " | " + code + "]");
      values[r][0] = platform;
      values[r][1] = title;
      values[r][2] = code;
      changed++;
    }
  }

  if (changed === 0) {
    Logger.log("cleanupSignupOriginColumns: nothing to clean.");
    return;
  }

  range.setValues(values);
  Logger.log("cleanupSignupOriginColumns: cleaned " + changed + " row(s).");
}

/** Keeps the left-most comma-separated part, trimmed. Non-strings pass through. */
function cleanLeftmost_(value) {
  if (typeof value !== "string" || value.indexOf(",") === -1) return value;
  return value.split(",")[0].trim();
}

/**
 * Cleans a possibly-conjoined event title. Titles themselves can contain
 * commas, so: keep the value as-is when it already IS a known title; else keep
 * the LONGEST comma-joined prefix that matches a known title; else fall back
 * to the plain first comma-part.
 */
function cleanTitleKeepingKnown_(value, knownTitles) {
  if (typeof value !== "string" || value.indexOf(",") === -1) return value;

  var norm = normalizeByStrippingWhiteSpaceAtTheEnd(value);
  if (norm && knownTitles.has(norm)) return value.trim();

  var parts = value.split(",");
  for (var take = parts.length - 1; take >= 1; take--) {
    var candidate = parts.slice(0, take).join(",").trim();
    var candNorm = normalizeByStrippingWhiteSpaceAtTheEnd(candidate);
    if (candNorm && knownTitles.has(candNorm)) return candidate;
  }
  return parts[0].trim();
}
