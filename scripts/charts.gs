/**
 * Robust number extractor for cells which might be:
 * - a real number,
 * - a string like "TTL RSVP =57" or "TTL ATTND =0",
 * - or a string with commas/decimal points.
 */
function extractNumberFromCell(cell) {
  if (cell === null || cell === undefined || cell === "") return 0;
  if (typeof cell === "number") return cell;
  const s = String(cell).trim();
  const numMatch = s.match(/-?\d{1,3}(?:,\d{3})*(?:\.\d+)?|-?\d+(?:\.\d+)?/);
  if (!numMatch) return 0;
  const rawNum = numMatch[0].replace(/,/g, "");
  const parsed = Number(rawNum);
  return isNaN(parsed) ? 0 : parsed;
}

/**
 * Extract event data from Contact List sheet.
 * Returns array of event objects with all relevant data.
 */
function extractEventData() {
  const [contactListSheet] = sheetsByName();

  const DATE_ROW      = ROW_NUMBERS.ROW_6;
  const EVENT_NAME_ROW = ROW_NUMBERS.ROW_7;
  const RSVP_ROW      = ROW_NUMBERS.ROW_9;
  const ATTENDED_ROW  = ROW_NUMBERS.ROW_10;
  const START_COL     = HELPER_CONSTANTS.EVENT_NAMES_START_COL;

  const lastCol = contactListSheet.getLastColumn();
  const numCols = lastCol - START_COL + 1;
  if (numCols <= 0) throw new Error("No event columns found starting at configured START_COL.");

  const dates      = contactListSheet.getRange(DATE_ROW,       START_COL, 1, numCols).getValues()[0];
  const eventNames = contactListSheet.getRange(EVENT_NAME_ROW, START_COL, 1, numCols).getValues()[0];
  const rsvpRaw    = contactListSheet.getRange(RSVP_ROW,       START_COL, 1, numCols).getValues()[0];
  const attendedRaw= contactListSheet.getRange(ATTENDED_ROW,   START_COL, 1, numCols).getValues()[0];

  const events = [];
  for (let i = 0; i < eventNames.length; i++) {
    const name = eventNames[i];
    if (!name) continue;

    const dateCell  = dates[i];
    const rsvpNum   = extractNumberFromCell(rsvpRaw[i]);
    const attendNum = extractNumberFromCell(attendedRaw[i]);

    events.push({
      name:       name,
      date:       dateCell instanceof Date ? dateCell : String(dateCell),
      rsvp:       rsvpNum,
      attended:   attendNum,
      difference: Math.abs(rsvpNum - attendNum)
    });
  }

  return events
}

// ─── HELPERS ─────────────────────────────────────────────────────────────────

function getYearFromDate_(dateVal) {
  if (dateVal instanceof Date && !isNaN(dateVal)) return dateVal.getFullYear();
  if (dateVal) {
    const d = new Date(dateVal);
    if (!isNaN(d)) return d.getFullYear();
  }
  return null;
}

function groupEventsByYear_(events) {
  const byYear = {};
  events.forEach(e => {
    const year = getYearFromDate_(e.date);
    if (!year) return;
    if (!byYear[year]) byYear[year] = [];
    byYear[year].push(e);
  });
  return byYear;
}

function computeStats_(events) {
  const total       = events.length;
  const totalRSVP   = events.reduce((s, e) => s + e.rsvp,     0);
  const totalAtt    = events.reduce((s, e) => s + e.attended,  0);
  const avgRSVP     = total > 0 ? totalRSVP / total  : 0;
  const avgAtt      = total > 0 ? totalAtt  / total  : 0;
  const attRate     = totalRSVP > 0 ? (totalAtt / totalRSVP * 100) : 0;
  const noShowRate  = totalRSVP > 0 ? ((totalRSVP - totalAtt) / totalRSVP * 100) : 0;
  return { total, totalRSVP, totalAtt, avgRSVP, avgAtt, attRate, noShowRate };
}

function round1_(n) { return Math.round(n * 10) / 10; }

/**
 * Write a section title into a sheet cell and return the next row.
 */
function writeTitle_(sheet, row, col, text, fontSize) {
  const c = sheet.getRange(row, col);
  c.setValue(text);
  c.setFontSize(fontSize || 12).setFontWeight("bold");
  return row + 1;
}

/**
 * Clear all charts from a sheet.
 */
function clearCharts_(sheet) {
  sheet.getCharts().forEach(ch => sheet.removeChart(ch));
}

/**
 * Insert a horizontal BAR chart (best for long event names).
 * Names go on the Y-axis; left margin is widened so full names are visible.
 *
 * @param {Sheet}  sheet
 * @param {Range}  dataRange  - includes header row
 * @param {number} anchorRow  - row to position chart at
 * @param {number} anchorCol  - column to position chart at
 * @param {string} title
 * @param {string} color      - hex colour for the bar series
 * @param {number} numItems   - number of data rows (for dynamic height)
 * @param {string} xAxisLabel
 */
function insertBarChart_(sheet, dataRange, anchorRow, anchorCol, title, color, numItems, xAxisLabel) {
  SpreadsheetApp.flush(); // commit all pending cell writes before chart reads the range
  const chart = sheet.newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(dataRange)
    .setNumHeaders(1)
    .setPosition(anchorRow, anchorCol, 0, 0)
    .setOption("title", title)
    .setOption("hAxis", { title: xAxisLabel || "Count", minValue: 0 })
    .setOption("vAxis", { textStyle: { fontSize: 10 } })
    .setOption("series", { 0: { color: color } })
    .setOption("legend", { position: "none" })
    .build();
  sheet.insertChart(chart);
}

/**
 * Insert a LINE chart for a timeline (event name on X-axis).
 * Chart is made wide and bottom labels are slanted so full names show.
 */
function insertLineChart_(sheet, dataRange, anchorRow, anchorCol, title, numItems) {
  SpreadsheetApp.flush(); // commit all pending cell writes before chart reads the range
  const chart = sheet.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(dataRange)
    .setNumHeaders(1)
    .setPosition(anchorRow, anchorCol, 0, 0)
    .setOption("title", title)
    .setOption("vAxis", { title: "Count", minValue: 0 })
    .setOption("hAxis", { title: "Event", slantedText: true, slantedTextAngle: 60, textStyle: { fontSize: 9 } })
    .setOption("series", {
      0: { color: "#4285F4", lineWidth: 2, labelInLegend: "RSVP" },
      1: { color: "#34A853", lineWidth: 2, labelInLegend: "Attended" }
    })
    .setOption("legend", { position: "bottom" })
    .setOption("curveType", "function")
    .build();
  sheet.insertChart(chart);
}

// ─── OVERVIEW SHEET SECTIONS ─────────────────────────────────────────────────

/**
 * Year-by-Year Summary Table + Column Chart
 * Writes to the provided sheet starting at startRow.
 * Returns the next available row after all content.
 */
function createYearlySummarySection_(sheet, startRow) {
  const events = extractEventData();
  const byYear = groupEventsByYear_(events);
  const years  = Object.keys(byYear).map(Number).sort();

  if (years.length === 0) {
    sheet.getRange(startRow, 1).setValue("No event data with recognisable dates found.");
    return startRow + 2;
  }

  startRow = writeTitle_(sheet, startRow, 1, "Year-by-Year Summary", 13);

  // ── Summary table ──────────────────────────────────────────────────────────
  const tblHeaders = [
    "Year", "# Events",
    "Total RSVPs", "Total Attended",
    "Avg RSVPs / Event", "Avg Attended / Event",
    "Attendance Rate (%)", "No-Show Rate (%)"
  ];
  const hdrRange = sheet.getRange(startRow, 1, 1, tblHeaders.length);
  hdrRange.setValues([tblHeaders]).setFontWeight("bold").setBackground("#D8E4BC");
  startRow++;

  const tableDataStartRow = startRow;
  const tableRows = [];
  years.forEach(year => {
    const s = computeStats_(byYear[year]);
    tableRows.push([
      year,
      s.total,
      s.totalRSVP,
      s.totalAtt,
      round1_(s.avgRSVP),
      round1_(s.avgAtt),
      round1_(s.attRate),
      round1_(s.noShowRate)
    ]);
  });

  // Totals / averages row
  const all = computeStats_(events);
  tableRows.push([
    "ALL YEARS",
    all.total,
    all.totalRSVP,
    all.totalAtt,
    round1_(all.avgRSVP),
    round1_(all.avgAtt),
    round1_(all.attRate),
    round1_(all.noShowRate)
  ]);

  sheet.getRange(tableDataStartRow, 1, tableRows.length, tblHeaders.length).setValues(tableRows);
  // Bold + highlight totals row
  sheet.getRange(tableDataStartRow + tableRows.length - 1, 1, 1, tblHeaders.length)
       .setFontWeight("bold").setBackground("#E8F0FE");

  // ── Column chart: RSVPs & Attended by Year ─────────────────────────────────
  // Write chart data off to the right (col 10) so it doesn't overlap the table
  const chartDataCol   = 10;
  const chartDataStart = tableDataStartRow - 1; // include header
  const chartData = [["Year", "Total RSVPs", "Total Attended"]];
  years.forEach(year => {
    const s = computeStats_(byYear[year]);
    chartData.push([String(year), s.totalRSVP, s.totalAtt]);
  });
  sheet.getRange(chartDataStart, chartDataCol, chartData.length, 3).setValues(chartData);
  SpreadsheetApp.flush();

  const compChart = sheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(sheet.getRange(chartDataStart, chartDataCol, chartData.length, 3))
    .setNumHeaders(1)
    .setPosition(tableDataStartRow, chartDataCol + 4, 0, 0)
    .setOption("title", "RSVPs & Attendance by Year")
    .setOption("vAxis", { title: "Count", minValue: 0 })
    .setOption("hAxis", { title: "Year" })
    .setOption("series", {
      0: { color: "#4285F4", labelInLegend: "Total RSVPs" },
      1: { color: "#34A853", labelInLegend: "Total Attended" }
    })
    .setOption("legend", { position: "bottom" })
    .build();
  sheet.insertChart(compChart);

  // ── Line chart: Avg RSVP & Avg Attended trend ─────────────────────────────
  const avgDataCol = chartDataCol;
  const avgDataStart = chartDataStart + chartData.length + 2;
  const avgData = [["Year", "Avg RSVPs / Event", "Avg Attended / Event"]];
  years.forEach(year => {
    const s = computeStats_(byYear[year]);
    avgData.push([String(year), round1_(s.avgRSVP), round1_(s.avgAtt)]);
  });
  sheet.getRange(avgDataStart, avgDataCol, avgData.length, 3).setValues(avgData);
  SpreadsheetApp.flush();

  const avgChart = sheet.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(sheet.getRange(avgDataStart, avgDataCol, avgData.length, 3))
    .setNumHeaders(1)
    .setPosition(tableDataStartRow + 22, chartDataCol + 4, 0, 0)
    .setOption("title", "Average RSVPs & Attendance per Event by Year")
    .setOption("vAxis", { title: "Count", minValue: 0 })
    .setOption("hAxis", { title: "Year" })
    .setOption("series", {
      0: { color: "#4285F4", lineWidth: 2, pointSize: 6, labelInLegend: "Avg RSVPs / Event" },
      1: { color: "#34A853", lineWidth: 2, pointSize: 6, labelInLegend: "Avg Attended / Event" }
    })
    .setOption("legend", { position: "bottom" })
    .setOption("curveType", "function")
    .build();
  sheet.insertChart(avgChart);

  return tableDataStartRow + tableRows.length + 2;
}

/**
 * Frequent Attendees section (3+ events attended), placed on the overview sheet.
 */
function createFrequentAttendeesSection_(sheet, startRow) {
  const [contactListSheet] = sheetsByName();
  const lastCol = contactListSheet.getLastColumn();
  const lastRow = contactListSheet.getLastRow();

  const headerRowValues = contactListSheet.getRange(ROW_NUMBERS.ROW_5, 1, 1, lastCol).getValues()[0];
  const attendedColIndex0 = headerRowValues.indexOf(COL_CONSTANTS.EVENTS_ATTENDED);
  if (attendedColIndex0 === -1) {
    sheet.getRange(startRow, 1).setValue("Could not find '# Events Attended' column.");
    return startRow + 2;
  }
  const attendedCol1 = attendedColIndex0 + 1;

  const dataStartRow = ROW_NUMBERS.ROW_12;
  const numRows = lastRow - dataStartRow + 1;
  if (numRows <= 0) {
    sheet.getRange(startRow, 1).setValue("No data rows found in Contact List.");
    return startRow + 2;
  }

  const namesValues    = contactListSheet.getRange(dataStartRow, 1,            numRows, 1).getValues();
  const attendedValues = contactListSheet.getRange(dataStartRow, attendedCol1, numRows, 1).getValues();

  const seen = new Set();
  const frequentAttendees = [];
  for (let i = 0; i < namesValues.length; i++) {
    const name = String(namesValues[i][0]).trim();
    if (!name || seen.has(name)) continue;
    const attended = typeof attendedValues[i][0] === "number"
      ? attendedValues[i][0]
      : Number(attendedValues[i][0]);
    if (isNaN(attended) || attended < 3) continue;
    seen.add(name);
    frequentAttendees.push([name, attended]);
  }

  const titleCell = sheet.getRange(startRow, 1);
  titleCell.setValue("Frequent Attendees (3+ Events)");
  titleCell.setFontSize(12).setFontWeight("bold");

  if (frequentAttendees.length === 0) {
    sheet.getRange(startRow + 1, 1).setValue("No attendees have attended 3 or more events yet.");
    return startRow + 3;
  }

  frequentAttendees.sort((a, b) => b[1] - a[1]);

  const hdrRange = sheet.getRange(startRow + 1, 1, 1, 2);
  hdrRange.setValues([["Name", "Times Attended"]]).setFontWeight("bold");

  const dataRange = sheet.getRange(startRow + 2, 1, frequentAttendees.length, 2);
  dataRange.setValues(frequentAttendees);
  dataRange.offset(0, 1, frequentAttendees.length, 1).setNumberFormat("0");

  Logger.log("Frequent attendees: " + frequentAttendees.length + " people attended 3+ times");
  return startRow + 2 + frequentAttendees.length + 2;
}

// ─── PER-YEAR SHEET SECTIONS ──────────────────────────────────────────────────

/**
 * Top-N RSVP'd events bar chart for a given year.
 */
function createYearTopRSVPChart_(sheet, year, events, startRow) {
  const sorted = events
    .filter(e => e.rsvp > 0)
    .sort((a, b) => b.rsvp - a.rsvp)
    .slice(0, 20);

  if (sorted.length === 0) {
    sheet.getRange(startRow, 1).setValue("No RSVP data for " + year);
    return startRow + 2;
  }

  const headers  = ["Event", "RSVP Count"];
  const dataRows = sorted.map(e => [e.name, e.rsvp]);
  const data     = [headers, ...dataRows];

  sheet.getRange(startRow, 1, data.length, 2).setValues(data);
  sheet.getRange(startRow, 1, 1, 2).setFontWeight("bold").setBackground("#D8E4BC");
  sheet.getRange(startRow + 1, 2, dataRows.length, 1).setNumberFormat("0");

  insertBarChart_(
    sheet,
    sheet.getRange(startRow, 1, data.length, 2),
    startRow, 4,
    "Top RSVP'd Events — " + year,
    "#4285F4",
    sorted.length,
    "RSVP Count"
  );

  return startRow + data.length + 2;
}

/**
 * Top-N Attended events bar chart for a given year.
 */
function createYearTopAttendedChart_(sheet, year, events, startRow) {
  const sorted = events
    .filter(e => e.attended > 0)
    .sort((a, b) => b.attended - a.attended)
    .slice(0, 20);

  if (sorted.length === 0) {
    sheet.getRange(startRow, 1).setValue("No attendance data for " + year);
    return startRow + 2;
  }

  const headers  = ["Event", "Attended"];
  const dataRows = sorted.map(e => [e.name, e.attended]);
  const data     = [headers, ...dataRows];

  sheet.getRange(startRow, 1, data.length, 2).setValues(data);
  sheet.getRange(startRow, 1, 1, 2).setFontWeight("bold").setBackground("#D8E4BC");
  sheet.getRange(startRow + 1, 2, dataRows.length, 1).setNumberFormat("0");

  insertBarChart_(
    sheet,
    sheet.getRange(startRow, 1, data.length, 2),
    startRow, 4,
    "Top Attended Events — " + year,
    "#34A853",
    sorted.length,
    "Attended Count"
  );

  return startRow + data.length + 2;
}

/**
 * Timeline (RSVP vs Attended line chart) for a given year, sorted by date.
 */
function createYearTimelineChart_(sheet, year, events, startRow) {
  const sorted = events
    .filter(e => e.rsvp > 0 || e.attended > 0)
    .sort((a, b) => {
      const da = a.date instanceof Date ? a.date : new Date(a.date);
      const db = b.date instanceof Date ? b.date : new Date(b.date);
      return da - db;
    });

  if (sorted.length === 0) {
    sheet.getRange(startRow, 1).setValue("No event data for " + year);
    return startRow + 2;
  }

  const headers  = ["Event", "RSVP", "Attended"];
  const dataRows = sorted.map(e => [e.name, e.rsvp, e.attended]);
  const data     = [headers, ...dataRows];

  sheet.getRange(startRow, 1, data.length, 3).setValues(data);
  sheet.getRange(startRow, 1, 1, 3).setFontWeight("bold").setBackground("#D8E4BC");
  sheet.getRange(startRow + 1, 2, dataRows.length, 2).setNumberFormat("0");

  insertLineChart_(
    sheet,
    sheet.getRange(startRow, 1, data.length, 3),
    startRow, 5,
    "RSVP vs Attended Timeline — " + year,
    sorted.length
  );

  return startRow + data.length + 2;
}

/**
 * Stats summary box for a given year (written as a small table).
 */
function createYearStatsSummary_(sheet, year, events, startRow) {
  const s = computeStats_(events);

  const rows = [
    ["Stat", "Value"],
    ["Year",                       year],
    ["# Events",                   s.total],
    ["Total RSVPs",                s.totalRSVP],
    ["Total Attended",             s.totalAtt],
    ["Avg RSVPs / Event",          round1_(s.avgRSVP)],
    ["Avg Attended / Event",       round1_(s.avgAtt)],
    ["Attendance Rate (%)",        round1_(s.attRate)],
    ["No-Show Rate (%)",           round1_(s.noShowRate)],
    ["Highest RSVP Event",         events.filter(e=>e.rsvp>0).sort((a,b)=>b.rsvp-a.rsvp)[0]?.name || "N/A"],
    ["Highest Attended Event",     events.filter(e=>e.attended>0).sort((a,b)=>b.attended-a.attended)[0]?.name || "N/A"]
  ];

  sheet.getRange(startRow, 1, rows.length, 2).setValues(rows);
  sheet.getRange(startRow, 1, 1, 2).setFontWeight("bold").setBackground("#D8E4BC");
  sheet.getRange(startRow, 1, rows.length, 2)
       .setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);

  // Auto-resize col A to show full event names
  sheet.autoResizeColumn(1);

  return startRow + rows.length + 2;
}

/**
 * Write all per-year sections (stats, timeline, top charts) into a single sheet,
 * stacked vertically, separated by a divider row.
 */
function createAllYearSections_(sheet, startRow) {
  const allEvents = extractEventData();
  const byYear    = groupEventsByYear_(allEvents);
  const years     = Object.keys(byYear).map(Number).sort();

  years.forEach(year => {
    const events = byYear[year];

    // Year divider header
    const dividerCell = sheet.getRange(startRow, 1);
    dividerCell.setValue("━━━  " + year + "  ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
    dividerCell.setFontSize(13).setFontWeight("bold").setFontColor("#FFFFFF").setBackground("#3C4043");
    startRow += 2;

    // Stats summary
    startRow = writeTitle_(sheet, startRow, 1, "Key Statistics — " + year, 12);
    startRow = createYearStatsSummary_(sheet, year, events, startRow);

    // Timeline
    startRow = writeTitle_(sheet, startRow, 1, "Timeline: RSVP vs Attended — " + year, 12);
    startRow = createYearTimelineChart_(sheet, year, events, startRow);

    // Top RSVP'd
    startRow = writeTitle_(sheet, startRow, 1, "Top RSVP'd Events (up to 20) — " + year, 12);
    startRow = createYearTopRSVPChart_(sheet, year, events, startRow);

    // Top Attended
    startRow = writeTitle_(sheet, startRow, 1, "Top Attended Events (up to 20) — " + year, 12);
    startRow = createYearTopAttendedChart_(sheet, year, events, startRow);

    startRow += 2; // breathing room between years
    Logger.log("Added year section: " + year + " (" + events.length + " events)");
  });

  return startRow;
}

// ─── MAIN ENTRY POINT ────────────────────────────────────────────────────────

/**
 * Main function — creates / refreshes a single "Data Analysis Graphs" sheet
 * containing: all-years summary → frequent attendees → per-year sections.
 */
function createRSVPvsAttendanceChart() {
  const [contactListSheet] = sheetsByName();
  const ss = contactListSheet.getParent();

  const sheetName = "Data Analysis Graphs";
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  } else {
    sheet.clear();
    clearCharts_(sheet);
  }

  sheet.setColumnWidth(1, 280);

  let row = 1;
  row = writeTitle_(sheet, row, 1, "RSVP & Attendance Analysis Dashboard", 14);
  row++;

  // All-years summary table + charts
  row = createYearlySummarySection_(sheet, row);

  // Frequent attendees
  row = createFrequentAttendeesSection_(sheet, row);
  row += 2;

  // Per-year sections (all in same sheet)
  row = createAllYearSections_(sheet, row);

  Logger.log("Dashboard complete in single sheet: " + sheetName);
}
