/**
 * Robust number extractor for cells which might be:
 * - a real number,
 * - a string like "TTL RSVP =57" or "TTL ATTND =0",
 * - or a string with commas/decimal points.
 */
function extractNumberFromCell(cell) {
  if (cell === null || cell === undefined || cell === "") return 0;

  // If it's already a number, return it
  if (typeof cell === "number") return cell;

  // Convert to string and trim whitespace
  const s = String(cell).trim();

  // Regex to find numbers like:
  //  123, 1,234, -123, 123.45, 1,234.56
  const numMatch = s.match(/-?\d{1,3}(?:,\d{3})*(?:\.\d+)?|-?\d+(?:\.\d+)?/);
  if (!numMatch) return 0;

  // Remove commas and parse
  const rawNum = numMatch[0].replace(/,/g, "");
  const parsed = Number(rawNum);
  return isNaN(parsed) ? 0 : parsed;
}

/**
 * Extract event data from Contact List sheet
 * Returns array of event objects with all relevant data
 */
function extractEventData() {
  const [contactListSheet] = sheetsByName();

  // Constants
  const DATE_ROW = ROW_NUMBERS.ROW_6;
  const EVENT_NAME_ROW = ROW_NUMBERS.ROW_7;
  const RSVP_ROW = ROW_NUMBERS.ROW_9;
  const ATTENDED_ROW = ROW_NUMBERS.ROW_10;
  const START_COL = HELPER_CONSTANTS.EVENT_NAMES_START_COL;

  // Determine how many event columns to read
  const lastCol = contactListSheet.getLastColumn();
  const numCols = lastCol - START_COL + 1;
  if (numCols <= 0) throw new Error("No event columns found starting at configured START_COL.");

  // Read ranges
  const dates = contactListSheet.getRange(DATE_ROW, START_COL, 1, numCols).getValues()[0];
  const eventNames = contactListSheet.getRange(EVENT_NAME_ROW, START_COL, 1, numCols).getValues()[0];
  const rsvpRaw = contactListSheet.getRange(RSVP_ROW, START_COL, 1, numCols).getValues()[0];
  const attendedRaw = contactListSheet.getRange(ATTENDED_ROW, START_COL, 1, numCols).getValues()[0];

  // Build event objects
  const events = [];
  for (let i = 0; i < eventNames.length; i++) {
    const name = eventNames[i];
    if (!name) continue; // Skip empty event names

    const dateCell = dates[i];
    const rsvpNum = extractNumberFromCell(rsvpRaw[i]);
    const attendNum = extractNumberFromCell(attendedRaw[i]);

    events.push({
      name: name,
      date: dateCell instanceof Date ? dateCell : String(dateCell),
      rsvp: rsvpNum,
      attended: attendNum,
      difference: Math.abs(rsvpNum - attendNum)
    });
  }

  return events;
}

/**
 * Format a date for display — returns "MM/DD/YYYY" string or the raw value
 */
function formatDateForDisplay_(dateVal) {
  if (dateVal instanceof Date && !isNaN(dateVal)) {
    return (dateVal.getMonth() + 1) + "/" + dateVal.getDate() + "/" + dateVal.getFullYear();
  }
  return String(dateVal);
}

/**
 * Section 1: Top RSVP'd Events
 * Bar chart showing events ranked by RSVP count (highest first)
 */
function createTopRSVPChart(chartSheet, startRow) {
  const events = extractEventData();

  const sorted = events
    .filter(e => e.rsvp > 0)
    .sort((a, b) => b.rsvp - a.rsvp)
    .slice(0, 20);

  if (sorted.length === 0) {
    chartSheet.getRange(startRow, 1).setValue("No events with RSVPs found");
    return startRow + 2;
  }

  const headers = ["Event", "RSVP Count", "Date"];
  const dataRows = sorted.map(e => [
    e.name,
    e.rsvp,
    e.date
  ]);
  const data = [headers, ...dataRows];

  chartSheet.getRange(startRow, 1, data.length, headers.length).setValues(data);
  chartSheet.getRange(startRow + 1, 2, dataRows.length, 1).setNumberFormat("0");
  if (sorted[0].date instanceof Date) {
    chartSheet.getRange(startRow + 1, 3, dataRows.length, 1).setNumberFormat("mm/dd/yyyy");
  }

  const chart = chartSheet.newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(chartSheet.getRange(startRow, 1, data.length, 2))
    .setPosition(startRow, 5, 0, 0)
    .setOption("title", "Top RSVP'd Events")
    .setOption("hAxis", { title: "RSVP Count" })
    .setOption("series", { 0: { color: "#4285F4" } })
    .setOption("legend", { position: "none" })
    .build();
  chartSheet.insertChart(chart);

  return startRow + data.length + 2;
}

/**
 * Section 2: Top Attended Events
 * Bar chart showing events ranked by attendance count (highest first)
 */
function createTopAttendedChart(chartSheet, startRow) {
  const events = extractEventData();

  const sorted = events
    .filter(e => e.attended > 0)
    .sort((a, b) => b.attended - a.attended)
    .slice(0, 20);

  if (sorted.length === 0) {
    chartSheet.getRange(startRow, 1).setValue("No events with attendance data found");
    return startRow + 2;
  }

  const headers = ["Event", "Attended", "Date"];
  const dataRows = sorted.map(e => [
    e.name,
    e.attended,
    e.date
  ]);
  const data = [headers, ...dataRows];

  chartSheet.getRange(startRow, 1, data.length, headers.length).setValues(data);
  chartSheet.getRange(startRow + 1, 2, dataRows.length, 1).setNumberFormat("0");
  if (sorted[0].date instanceof Date) {
    chartSheet.getRange(startRow + 1, 3, dataRows.length, 1).setNumberFormat("mm/dd/yyyy");
  }

  const chart = chartSheet.newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(chartSheet.getRange(startRow, 1, data.length, 2))
    .setPosition(startRow, 5, 0, 0)
    .setOption("title", "Top Attended Events")
    .setOption("hAxis", { title: "Attended Count" })
    .setOption("series", { 0: { color: "#34A853" } })
    .setOption("legend", { position: "none" })
    .build();
  chartSheet.insertChart(chart);

  return startRow + data.length + 2;
}

/**
 * Section 3: All Events Timeline sorted by date
 * Line graph showing RSVP vs Attended over time
 */
function createAllEventsTimelineChart(chartSheet, startRow) {
  const events = extractEventData();

  const sortedEvents = events
    .filter(e => e.rsvp > 0)
    .sort((a, b) => {
      const dateA = a.date instanceof Date ? a.date : new Date(a.date);
      const dateB = b.date instanceof Date ? b.date : new Date(b.date);
      return dateA - dateB;
    });

  if (sortedEvents.length === 0) {
    chartSheet.getRange(startRow, 1).setValue("No events with RSVPs found");
    return startRow + 2;
  }

  const headers = ["Event Name", "RSVP", "Attended", "Date"];
  const dataRows = sortedEvents.map(e => [e.name, e.rsvp, e.attended, e.date]);
  const data = [headers, ...dataRows];

  chartSheet.getRange(startRow, 1, data.length, headers.length).setValues(data);
  chartSheet.getRange(startRow + 1, 2, dataRows.length, 2).setNumberFormat("0");
  if (sortedEvents[0].date instanceof Date) {
    chartSheet.getRange(startRow + 1, 4, dataRows.length, 1).setNumberFormat("mm/dd/yyyy");
  }

  const chartRange = chartSheet.getRange(startRow, 1, data.length, 3);
  const chart = chartSheet.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(chartRange)
    .setPosition(startRow, 6, 0, 0)
    .setOption("title", "All Events Timeline (RSVP vs Attended)")
    .setOption("vAxis", { title: "Count" })
    .setOption("hAxis", { title: "Event", slantedText: true, slantedTextAngle: 45 })
    .setOption("series", {
      0: { color: "#4285F4", lineWidth: 2, labelInLegend: "RSVP" },
      1: { color: "#34A853", lineWidth: 2, labelInLegend: "Attended" }
    })
    .setOption("legend", { position: "bottom" })
    .setOption("curveType", "function")
    .build();
  chartSheet.insertChart(chart);

  return startRow + data.length + 2;
}

/**
 * Section 4: Overall RSVP vs Attended summary stats
 */
function createRSVPvsAttendedComparisonChart(chartSheet, startRow) {
  const events = extractEventData();

  const eventsWithRSVP = events.filter(e => e.rsvp > 0);

  if (eventsWithRSVP.length === 0) {
    chartSheet.getRange(startRow, 1).setValue("No events with RSVPs found");
    return startRow + 2;
  }

  const totalRSVP = eventsWithRSVP.reduce((sum, e) => sum + e.rsvp, 0);
  const totalAttended = eventsWithRSVP.reduce((sum, e) => sum + e.attended, 0);
  const avgRSVP = totalRSVP / eventsWithRSVP.length;
  const avgAttended = totalAttended / eventsWithRSVP.length;
  const attendanceRate = totalRSVP > 0 ? (totalAttended / totalRSVP * 100) : 0;

  const headers = ["Metric", "RSVP", "Attended"];
  const data = [
    headers,
    ["Total", totalRSVP, totalAttended],
    ["Average per Event", avgRSVP, avgAttended],
    ["Event Count", eventsWithRSVP.length, eventsWithRSVP.length],
    ["Attendance Rate (%)", "", attendanceRate]
  ];

  chartSheet.getRange(startRow, 1, data.length, headers.length).setValues(data);
  chartSheet.getRange(startRow + 1, 2, 3, 2).setNumberFormat("0.00");
  chartSheet.getRange(startRow + 4, 3, 1, 1).setNumberFormat("0.00");

  const chartRange = chartSheet.getRange(startRow, 1, 3, 3);
  const chart = chartSheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(chartRange)
    .setPosition(startRow, 5, 0, 0)
    .setOption("title", "RSVP vs Actual Attendance")
    .setOption("vAxis", { title: "Count" })
    .setOption("series", {
      0: { color: "#4285F4", labelInLegend: "RSVP" },
      1: { color: "#34A853", labelInLegend: "Attended" }
    })
    .setOption("legend", { position: "bottom" })
    .build();
  chartSheet.insertChart(chart);

  return startRow + data.length + 2;
}

/**
 * Section 5: Frequent Attendees (attended 3+ times)
 */
function createFrequentAttendeesSection(chartSheet, startRow) {
  const [contactListSheet] = sheetsByName();

  const lastCol = contactListSheet.getLastColumn();
  const lastRow = contactListSheet.getLastRow();

  const headerRowValues = contactListSheet.getRange(ROW_NUMBERS.ROW_5, 1, 1, lastCol).getValues()[0];
  const attendedColIndex0 = headerRowValues.indexOf(COL_CONSTANTS.EVENTS_ATTENDED);
  if (attendedColIndex0 === -1) {
    chartSheet.getRange(startRow, 1).setValue("Could not find '# Events Attended' column.");
    return startRow + 2;
  }
  const attendedCol1 = attendedColIndex0 + 1;

  const dataStartRow = ROW_NUMBERS.ROW_12;
  const numRows = lastRow - dataStartRow + 1;
  if (numRows <= 0) {
    chartSheet.getRange(startRow, 1).setValue("No data rows found in Contact List.");
    return startRow + 2;
  }

  const namesValues    = contactListSheet.getRange(dataStartRow, 1,           numRows, 1).getValues();
  const attendedValues = contactListSheet.getRange(dataStartRow, attendedCol1, numRows, 1).getValues();

  const seen = new Set();
  const frequentAttendees = [];

  for (let i = 0; i < namesValues.length; i++) {
    const name = String(namesValues[i][0]).trim();
    if (!name) continue;

    const attendedRaw = attendedValues[i][0];
    const attended = typeof attendedRaw === "number" ? attendedRaw : Number(attendedRaw);
    if (isNaN(attended) || attended < 3) continue;

    if (seen.has(name)) continue;
    seen.add(name);
    frequentAttendees.push([name, attended]);
  }

  const titleCell = chartSheet.getRange(startRow, 1);
  titleCell.setValue("Frequent Attendees (Attended 3 or More Times)");
  titleCell.setFontSize(12).setFontWeight("bold");

  if (frequentAttendees.length === 0) {
    chartSheet.getRange(startRow + 1, 1).setValue("No attendees have attended 3 or more times yet.");
    return startRow + 3;
  }

  frequentAttendees.sort((a, b) => b[1] - a[1]);

  const headersRange = chartSheet.getRange(startRow + 1, 1, 1, 2);
  headersRange.setValues([["Name", "Times Attended"]]);
  headersRange.setFontWeight("bold");

  const dataRange = chartSheet.getRange(startRow + 2, 1, frequentAttendees.length, 2);
  dataRange.setValues(frequentAttendees);
  dataRange.offset(0, 1, frequentAttendees.length, 1).setNumberFormat("0");

  Logger.log("Frequent attendees section: " + frequentAttendees.length + " people attended 3 or more times");
  return startRow + 2 + frequentAttendees.length + 2;
}

/**
 * Main function to create all analysis charts
 */
function createRSVPvsAttendanceChart() {
  const [contactListSheet] = sheetsByName();
  const ss = contactListSheet.getParent();
  const chartSheetName = "Data Analysis Graphs";

  let chartSheet = ss.getSheetByName(chartSheetName);
  if (!chartSheet) {
    chartSheet = ss.insertSheet(chartSheetName);
  }

  chartSheet.clear();

  chartSheet.getRange(1, 1).setValue("RSVP & Attendance Analysis Dashboard");
  chartSheet.getRange(1, 1).setFontSize(14).setFontWeight("bold");

  let currentRow = 3;

  // Section 1: Top RSVP'd Events
  currentRow = createTopRSVPChart(chartSheet, currentRow);

  // Section 2: Top Attended Events
  currentRow = createTopAttendedChart(chartSheet, currentRow);

  // Section 3: All Events Timeline
  currentRow = createAllEventsTimelineChart(chartSheet, currentRow);

  // Section 4: Overall Comparison Stats
  currentRow = createRSVPvsAttendedComparisonChart(chartSheet, currentRow);

  // Section 5: Frequent Attendees
  currentRow = createFrequentAttendeesSection(chartSheet, currentRow);

  Logger.log("Dashboard created successfully in 'Data Analysis Graphs' sheet");
}
