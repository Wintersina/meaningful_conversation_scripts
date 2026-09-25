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
 * Top 10 RSVP'd events across ALL years (table + bar chart), placed on the
 * overview sheet. Each row shows the event name with its year so repeated
 * topics from different years stay distinguishable.
 */
function createAllYearsTopRSVPSection_(sheet, startRow) {
  const events = extractEventData();
  const top = events
    .filter(e => e.rsvp > 0)
    .sort((a, b) => b.rsvp - a.rsvp)
    .slice(0, 10);

  startRow = writeTitle_(sheet, startRow, 1, "Top 10 RSVP'd Events — All Years", 13);

  if (top.length === 0) {
    sheet.getRange(startRow, 1).setValue("No RSVP data found.");
    return startRow + 2;
  }

  const headers  = ["Event", "RSVP Count"];
  const dataRows = top.map(e => {
    const year = getYearFromDate_(e.date);
    return [e.name + (year ? " (" + year + ")" : ""), e.rsvp];
  });
  const data = [headers, ...dataRows];

  sheet.getRange(startRow, 1, data.length, 2).setValues(data);
  sheet.getRange(startRow, 1, 1, 2).setFontWeight("bold").setBackground("#D8E4BC");
  sheet.getRange(startRow + 1, 2, dataRows.length, 1).setNumberFormat("0");

  insertBarChart_(
    sheet,
    sheet.getRange(startRow, 1, data.length, 2),
    startRow, 4,
    "Top 10 RSVP'd Events — All Years",
    "#4285F4",
    top.length,
    "RSVP Count"
  );

  return startRow + data.length + 2;
}

/**
 * Frequent Attendees roster (3+ events attended), placed on the overview sheet.
 * Fed by the same per-person roll-up as the engagement sections above, so the
 * roster length always matches the "3+" row in the threshold table.
 */
function createFrequentAttendeesSection_(sheet, startRow, people) {
  people = people || extractPeopleEngagement_();

  const frequentAttendees = people
    .filter(p => p.attended >= FREQUENT_ATTENDEE_MIN)
    .sort((a, b) => b.attended - a.attended)
    .map(p => [p.name, p.attended]);

  const titleCell = sheet.getRange(startRow, 1);
  titleCell.setValue("Frequent Attendees (" + FREQUENT_ATTENDEE_MIN + "+ Events)");
  titleCell.setFontSize(12).setFontWeight("bold");

  if (frequentAttendees.length === 0) {
    sheet.getRange(startRow + 1, 1)
         .setValue("No attendees have attended " + FREQUENT_ATTENDEE_MIN + " or more events yet.");
    return startRow + 3;
  }

  sheet.getRange(startRow + 1, 1, 1, 2).setValues([["Name", "Times Attended"]]).setFontWeight("bold");

  const dataRange = sheet.getRange(startRow + 2, 1, frequentAttendees.length, 2);
  dataRange.setValues(frequentAttendees);
  dataRange.offset(0, 1, frequentAttendees.length, 1).setNumberFormat("0");

  Logger.log("Frequent attendees: " + frequentAttendees.length + " people attended " +
             FREQUENT_ATTENDEE_MIN + "+ times");
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

// ─── PERSON-CENTRIC ENGAGEMENT ───────────────────────────────────────────────
//
// Everything above this line counts EVENTS: the row-9/row-10 "TTL RSVP =" and
// "TTL ATTND =" cells are per-event totals, so somebody who RSVP'd to twelve
// events is twelve RSVPs up there. The sections below count PEOPLE instead,
// off the per-person "# Events RSVP'd" / "# Events Attended" counter columns,
// so one individual is one individual no matter how often they showed up.
// All of it is read-only against the Contact List.

/** Cumulative thresholds reported for both RSVPs and attendance. */
const ENGAGEMENT_THRESHOLDS = [1, 2, 3, 5, 10, 20, 50];

/** Non-overlapping frequency buckets, in display order (max null = no cap). */
const ENGAGEMENT_BUCKETS = [
  { label: "1 time",      min: 1,  max: 1    },
  { label: "2 times",     min: 2,  max: 2    },
  { label: "3–4 times",   min: 3,  max: 4    },
  { label: "5–9 times",   min: 5,  max: 9    },
  { label: "10–19 times", min: 10, max: 19   },
  { label: "20–49 times", min: 20, max: 49   },
  { label: "50+ times",   min: 50, max: null }
];

/** Rows an inserted chart covers, so the next section doesn't sit underneath it. */
const CHART_ROW_SPAN = 20;

/** Events attended before someone counts as a "frequent attendee". */
const FREQUENT_ATTENDEE_MIN = 3;

/** Longest outreach roster written before it's truncated with a "…and N more". */
const HIGH_INTEREST_LIST_LIMIT = 50;

/** Counter cells hold a number, or "" when their IFERROR fires. */
function toEngagementCount_(value) {
  if (typeof value === "number") return value > 0 ? value : 0;
  const n = Number(String(value).trim());
  return isNaN(n) || n < 0 ? 0 : n;
}

/**
 * Labels that are section furniture ("Stop RSVP", "Total RSVP'd", …) rather
 * than people. Built lazily: helpers.gs loads *after* charts.gs, so COL_CONSTANTS
 * must not be touched at file-evaluation time.
 */
function sectionLabelKeys_() {
  const keys = new Set();
  Object.keys(COL_CONSTANTS).forEach(k => {
    keys.add(normalizeByStrippingWhiteSpaceAtTheEnd(COL_CONSTANTS[k]));
  });
  return keys;
}

/**
 * Roll the Contact List up to one record per person.
 *
 * The sheet holds the same person more than once on purpose — the Attended and
 * RSVP 2+ sections at the top are copies of main-list rows — so rows are keyed
 * by the normalised col-A full-name key and the highest counter seen wins.
 *
 * @return {{name: string, rsvp: number, attended: number}[]}
 */
function extractPeopleEngagement_() {
  const [contactListSheet] = sheetsByName();
  const lastRow = contactListSheet.getLastRow();

  const rsvpCol     = findColMarker_(contactListSheet, MARKER_KEYS.EVENTS_RSVPD,    COL_CONSTANTS.EVENTS_RSVPD);
  const attendedCol = findColMarker_(contactListSheet, MARKER_KEYS.EVENTS_ATTENDED, COL_CONSTANTS.EVENTS_ATTENDED);
  if (rsvpCol === -1 || attendedCol === -1) {
    throw new Error('Could not find the "' + COL_CONSTANTS.EVENTS_RSVPD + '" and/or "' +
                    COL_CONSTANTS.EVENTS_ATTENDED + '" column in Row 5.');
  }

  const dataStartRow = ROW_NUMBERS.ROW_12;
  const numRows = lastRow - dataStartRow + 1;
  if (numRows <= 0) return [];

  const names    = contactListSheet.getRange(dataStartRow, 1,           numRows, 1).getValues();
  const rsvps    = contactListSheet.getRange(dataStartRow, rsvpCol,     numRows, 1).getValues();
  const attendeds= contactListSheet.getRange(dataStartRow, attendedCol, numRows, 1).getValues();

  const skip  = sectionLabelKeys_();
  const byKey = {};

  for (let i = 0; i < numRows; i++) {
    const display = String(names[i][0]).trim();
    if (!display) continue;

    const key = normalizeByStrippingWhiteSpaceAtTheEnd(display);
    if (!key || skip.has(key)) continue; // marker row, not a person

    const rsvp     = toEngagementCount_(rsvps[i][0]);
    const attended = toEngagementCount_(attendeds[i][0]);

    const existing = byKey[key];
    if (existing) {
      existing.rsvp     = Math.max(existing.rsvp,     rsvp);
      existing.attended = Math.max(existing.attended, attended);
    } else {
      byKey[key] = { name: display, rsvp: rsvp, attended: attended };
    }
  }

  const people = Object.keys(byKey).map(k => byKey[k]);
  Logger.log("Engagement roll-up: " + people.length + " unique individuals");
  return people;
}

/** How many people hit `field` at least `n` times. */
function countAtLeast_(people, field, n) {
  return people.filter(p => p[field] >= n).length;
}

/** [[bucketLabel, peopleInBucket], …] for one counter field. */
function bucketCounts_(people, field) {
  return ENGAGEMENT_BUCKETS.map(b => {
    const n = people.filter(p => {
      const v = p[field];
      return v >= b.min && (b.max === null || v <= b.max);
    }).length;
    return [b.label, n];
  });
}

function pct_(part, whole) {
  return whole > 0 ? round1_(part / whole * 100) : 0;
}

/**
 * Insert a COLUMN chart (vertical bars) — used where categories are short
 * labels and two series sit side by side.
 */
function insertColumnChart_(sheet, dataRange, anchorRow, anchorCol, title, hAxisTitle, series) {
  SpreadsheetApp.flush(); // commit all pending cell writes before chart reads the range
  const chart = sheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(dataRange)
    .setNumHeaders(1)
    .setPosition(anchorRow, anchorCol, 0, 0)
    .setOption("title", title)
    .setOption("vAxis", { title: "Individuals", minValue: 0 })
    .setOption("hAxis", { title: hAxisTitle || "" })
    .setOption("series", series)
    .setOption("legend", { position: "bottom" })
    .build();
  sheet.insertChart(chart);
}

/**
 * Headline head-count table: how many individuals RSVP'd, how many came, how
 * many did either more than once, and how the two populations overlap.
 */
function createEngagementOverviewSection_(sheet, startRow, people) {
  people = people || extractPeopleEngagement_();

  startRow = writeTitle_(sheet, startRow, 1, "Individual Engagement — People, Not Seats", 13);

  if (people.length === 0) {
    sheet.getRange(startRow, 1).setValue("No people found in the Contact List.");
    return startRow + 2;
  }

  const totalPeople   = people.length;
  const everRsvped    = countAtLeast_(people, "rsvp", 1);
  const rsvpedTwice   = countAtLeast_(people, "rsvp", 2);
  const everAttended  = countAtLeast_(people, "attended", 1);
  const attendedTwice = countAtLeast_(people, "attended", 2);
  const rsvpNoShow    = people.filter(p => p.rsvp >= 1 && p.attended === 0).length;
  const walkInOnly    = people.filter(p => p.attended >= 1 && p.rsvp === 0).length;
  const converted     = people.filter(p => p.rsvp >= 1 && p.attended >= 1).length;

  const rows = [
    ["Metric", "Individuals", "% of list"],
    ["Individuals in the Contact List",           totalPeople,   100],
    ["Ever RSVP'd",                               everRsvped,    pct_(everRsvped,    totalPeople)],
    ["RSVP'd more than once",                     rsvpedTwice,   pct_(rsvpedTwice,   totalPeople)],
    ["Ever attended",                             everAttended,  pct_(everAttended,  totalPeople)],
    ["Attended more than once",                   attendedTwice, pct_(attendedTwice, totalPeople)],
    ["RSVP'd but never attended",                 rsvpNoShow,    pct_(rsvpNoShow,    totalPeople)],
    ["Attended without ever RSVP'ing (walk-ins)", walkInOnly,    pct_(walkInOnly,    totalPeople)],
    ["RSVP'd and attended at least once",         converted,     pct_(converted,     totalPeople)]
  ];

  sheet.getRange(startRow, 1, rows.length, 3).setValues(rows);
  sheet.getRange(startRow, 1, 1, 3).setFontWeight("bold").setBackground("#D8E4BC");
  sheet.getRange(startRow + 1, 2, rows.length - 1, 1).setNumberFormat("0");
  sheet.getRange(startRow + 1, 3, rows.length - 1, 1).setNumberFormat("0.0");
  sheet.getRange(startRow, 1, rows.length, 3)
       .setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);

  let row = startRow + rows.length;
  sheet.getRange(row, 1).setValue(
    "Of the " + everRsvped + " individuals who ever RSVP'd, " + converted + " (" +
    pct_(converted, everRsvped) + "%) showed up at least once."
  );
  sheet.getRange(row, 1).setFontWeight("bold");
  row++;

  sheet.getRange(row, 1).setValue(
    "Head counts, de-duplicated by name — unlike \"Total RSVPs\" in the year table above, which counts one per " +
    "event. \"# Events RSVP'd\" counts only \"RSVP'd: yes\" (maybes excluded); \"# Events Attended\" counts every " +
    "\"Attended: yes\", walk-ins included."
  );
  sheet.getRange(row, 1).setFontStyle("italic").setFontSize(9);

  return row + 2;
}

/**
 * The cumulative view: how many individuals came (or RSVP'd) at least 1, 2, 3,
 * 5, 10, 20 and 50 times.
 */
function createEngagementThresholdSection_(sheet, startRow, people) {
  people = people || extractPeopleEngagement_();

  startRow = writeTitle_(sheet, startRow, 1, "Depth of Engagement — Individuals by Threshold", 13);

  if (people.length === 0) {
    sheet.getRange(startRow, 1).setValue("No people found in the Contact List.");
    return startRow + 2;
  }

  const headers  = ["Threshold", "Individuals RSVP'd", "Individuals Attended"];
  const dataRows = ENGAGEMENT_THRESHOLDS.map(n => [
    n + "+",
    countAtLeast_(people, "rsvp", n),
    countAtLeast_(people, "attended", n)
  ]);
  const data = [headers, ...dataRows];

  sheet.getRange(startRow, 1, data.length, 3).setValues(data);
  sheet.getRange(startRow, 1, 1, 3).setFontWeight("bold").setBackground("#D8E4BC");
  sheet.getRange(startRow + 1, 2, dataRows.length, 2).setNumberFormat("0");

  insertColumnChart_(
    sheet,
    sheet.getRange(startRow, 1, data.length, 3),
    startRow, 5,
    "Individuals Reaching Each Threshold",
    "Times RSVP'd / attended",
    {
      0: { color: "#4285F4", labelInLegend: "Individuals RSVP'd" },
      1: { color: "#34A853", labelInLegend: "Individuals Attended" }
    }
  );

  const noteRow = startRow + data.length;
  sheet.getRange(noteRow, 1).setValue("Cumulative: the \"5+\" row includes everyone in the 10+, 20+ and 50+ rows.");
  sheet.getRange(noteRow, 1).setFontStyle("italic").setFontSize(9);

  return startRow + Math.max(data.length + 1, CHART_ROW_SPAN) + 2;
}

/**
 * Shared body for the two frequency-distribution sections: non-overlapping
 * buckets (1, 2, 3–4, 5–9, 10–19, 20–49, 50+) plus a bar chart.
 */
function createFrequencyDistributionSection_(sheet, startRow, people, field, title, unitHeader, color) {
  people = people || extractPeopleEngagement_();

  startRow = writeTitle_(sheet, startRow, 1, title, 13);

  const counts  = bucketCounts_(people, field);
  const engaged = counts.reduce((s, r) => s + r[1], 0);
  if (engaged === 0) {
    sheet.getRange(startRow, 1).setValue("No " + unitHeader.toLowerCase() + " data found.");
    return startRow + 2;
  }

  const headers  = [unitHeader, "Individuals", "% of engaged"];
  const dataRows = counts.map(r => [r[0], r[1], pct_(r[1], engaged)]);
  const data     = [headers, ...dataRows];

  sheet.getRange(startRow, 1, data.length, 3).setValues(data);
  sheet.getRange(startRow, 1, 1, 3).setFontWeight("bold").setBackground("#D8E4BC");
  sheet.getRange(startRow + 1, 2, dataRows.length, 1).setNumberFormat("0");
  sheet.getRange(startRow + 1, 3, dataRows.length, 1).setNumberFormat("0.0");

  // Chart the label + count columns only — the % column stays in the table.
  insertBarChart_(
    sheet,
    sheet.getRange(startRow, 1, data.length, 2),
    startRow, 5,
    title,
    color,
    dataRows.length,
    "Individuals"
  );

  const noteRow = startRow + data.length;
  sheet.getRange(noteRow, 1).setValue(
    engaged + " individuals have at least one; the percentages are of those " + engaged + ", not of the whole list."
  );
  sheet.getRange(noteRow, 1).setFontStyle("italic").setFontSize(9);

  return startRow + Math.max(data.length + 1, CHART_ROW_SPAN) + 2;
}

function createAttendanceDistributionSection_(sheet, startRow, people) {
  return createFrequencyDistributionSection_(
    sheet, startRow, people, "attended",
    "Attendance Frequency — How Often Individuals Came",
    "Times Attended", "#34A853"
  );
}

function createRsvpDistributionSection_(sheet, startRow, people) {
  return createFrequencyDistributionSection_(
    sheet, startRow, people, "rsvp",
    "RSVP Frequency — How Often Individuals RSVP'd",
    "Times RSVP'd", "#4285F4"
  );
}

/**
 * The outreach roster: people who keep saying yes but have never walked in.
 * Highest declared interest, zero conversion — the shortest bridge to build.
 */
function createHighInterestNotConvertedSection_(sheet, startRow, people) {
  people = people || extractPeopleEngagement_();

  startRow = writeTitle_(sheet, startRow, 1,
    "High Interest, Not Yet Converted — RSVP'd 2+ Times, Never Attended", 13);

  const candidates = people
    .filter(p => p.rsvp >= 2 && p.attended === 0)
    .sort((a, b) => b.rsvp - a.rsvp);

  if (candidates.length === 0) {
    sheet.getRange(startRow, 1)
         .setValue("Nobody has RSVP'd twice or more without attending — everyone who keeps saying yes has shown up.");
    return startRow + 3;
  }

  const shown = candidates.slice(0, HIGH_INTEREST_LIST_LIMIT);
  const data  = [["Name", "Times RSVP'd"], ...shown.map(p => [p.name, p.rsvp])];

  sheet.getRange(startRow, 1, data.length, 2).setValues(data);
  sheet.getRange(startRow, 1, 1, 2).setFontWeight("bold").setBackground("#D8E4BC");
  sheet.getRange(startRow + 1, 2, shown.length, 1).setNumberFormat("0");

  let row = startRow + data.length;
  if (candidates.length > shown.length) {
    sheet.getRange(row, 1).setValue(
      "…and " + (candidates.length - shown.length) + " more (" + candidates.length + " in total)."
    );
    sheet.getRange(row, 1).setFontStyle("italic").setFontSize(9);
    row++;
  }

  Logger.log("High interest, not converted: " + candidates.length + " people");
  return row + 2;
}

// ─── MAIN ENTRY POINT ────────────────────────────────────────────────────────

/**
 * Main function — creates / refreshes a single "Data Analysis Graphs" sheet
 * containing, top to bottom: all-years event summary → top RSVP'd events →
 * person-centric engagement (head counts, depth thresholds, frequency
 * distributions, the not-yet-converted outreach roster) → frequent attendees →
 * per-year sections.
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

  // Top 10 RSVP'd events across all years
  row = createAllYearsTopRSVPSection_(sheet, row);
  row += 2;

  // Person-centric engagement. The Contact List is read once here and the same
  // roll-up is handed to every section below, including Frequent Attendees.
  const people = extractPeopleEngagement_();
  row = createEngagementOverviewSection_(sheet, row, people);
  row = createEngagementThresholdSection_(sheet, row, people);
  row = createAttendanceDistributionSection_(sheet, row, people);
  row = createRsvpDistributionSection_(sheet, row, people);
  row = createHighInterestNotConvertedSection_(sheet, row, people);
  row += 2;

  // Frequent attendees
  row = createFrequentAttendeesSection_(sheet, row, people);
  row += 2;

  // Per-year sections (all in same sheet)
  row = createAllYearSections_(sheet, row);

  Logger.log("Dashboard complete in single sheet: " + sheetName);
}
