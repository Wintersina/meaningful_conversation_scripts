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
 * Helper function to classify day type from a date
 */
function getDayType(dateCell) {
  if (dateCell instanceof Date && !isNaN(dateCell)) {
    const day = dateCell.getDay();
    return (day === 1) ? "Monday" : (day === 4) ? "Thursday" : "Other";
  } else {
    // Try to parse as date string
    const maybeDate = new Date(String(dateCell));
    if (!isNaN(maybeDate)) {
      const day = maybeDate.getDay();
      return (day === 1) ? "Monday" : (day === 4) ? "Thursday" : "Other";
    }
  }
  return "Unknown";
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
    const dayType = getDayType(dateCell);

    events.push({
      name: name,
      date: dateCell instanceof Date ? dateCell : String(dateCell),
      dayType: dayType,
      rsvp: rsvpNum,
      attended: attendNum,
      difference: Math.abs(rsvpNum - attendNum)
    });
  }

  return events;
}

/**
 * Graph 1: Average RSVP numbers by day (Monday vs Thursday)
 * Creates a column chart comparing average RSVPs
 */
function createAverageRSVPByDayChart(chartSheet, startRow) {
  const events = extractEventData();

  // Filter and calculate averages
  const mondayEvents = events.filter(e => e.dayType === "Monday");
  const thursdayEvents = events.filter(e => e.dayType === "Thursday");

  const mondayAvg = mondayEvents.length > 0
    ? mondayEvents.reduce((sum, e) => sum + e.rsvp, 0) / mondayEvents.length
    : 0;
  const thursdayAvg = thursdayEvents.length > 0
    ? thursdayEvents.reduce((sum, e) => sum + e.rsvp, 0) / thursdayEvents.length
    : 0;

  // Write data to sheet
  const headers = ["Day", "Average RSVP", "Event Count"];
  const data = [
    headers,
    ["Monday", mondayAvg, mondayEvents.length],
    ["Thursday", thursdayAvg, thursdayEvents.length]
  ];

  chartSheet.getRange(startRow, 1, data.length, headers.length).setValues(data);

  // Format numeric columns as numbers (not dates)
  chartSheet.getRange(startRow + 1, 2, data.length - 1, 2).setNumberFormat("0.00"); // Average RSVP and Event Count

  // Create chart - using separate series for better labeling
  const chart = chartSheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(chartSheet.getRange(startRow, 1, 1, 1)) // Day header
    .addRange(chartSheet.getRange(startRow + 1, 1, 2, 1)) // Day values
    .addRange(chartSheet.getRange(startRow, 2, 1, 1)) // Average RSVP header
    .addRange(chartSheet.getRange(startRow + 1, 2, 2, 1)) // Average RSVP values
    .setPosition(startRow, 5, 0, 0)
    .setOption("title", "Average RSVP: Monday vs Thursday")
    .setOption("vAxis", { title: "Average RSVP Count" })
    .setOption("hAxis", { title: "Day of Week" })
    .setOption("series", {
      0: { color: "#4285F4", labelInLegend: "Average RSVP" }
    })
    .setOption("legend", { position: "bottom" })
    .build();
  chartSheet.insertChart(chart);

  return startRow + data.length + 2; // Return next available row
}

/**
 * Graph 2: Events with closest RSVP to Attendance match (Top 15)
 * Line graph showing top 15 events where RSVP prediction was most accurate
 */
function createRSVPAccuracyChart(chartSheet, startRow) {
  const events = extractEventData();

  // Sort by smallest difference (best predictions)
  const sortedEvents = events
    .filter(e => e.rsvp > 0) // Only include events with RSVPs
    .sort((a, b) => a.difference - b.difference)
    .slice(0, 15); // Top 15 most accurate

  if (sortedEvents.length === 0) {
    chartSheet.getRange(startRow, 1).setValue("No events with RSVPs found");
    return startRow + 2;
  }

  // Write data to sheet - include date and day type
  const headers = ["Event Name", "RSVP", "Attended", "Difference", "Event Date", "Day Type"];
  const dataRows = sortedEvents.map(e => [
    e.name,
    e.rsvp,
    e.attended,
    e.difference,
    e.date,
    e.dayType
  ]);
  const data = [headers, ...dataRows];

  chartSheet.getRange(startRow, 1, data.length, headers.length).setValues(data);

  // Format numeric columns as numbers (not dates)
  chartSheet.getRange(startRow + 1, 2, dataRows.length, 3).setNumberFormat("0"); // RSVP, Attended, Difference
  // Format date column if they are Date objects
  if (sortedEvents[0].date instanceof Date) {
    chartSheet.getRange(startRow + 1, 5, dataRows.length, 1).setNumberFormat("mm/dd/yyyy");
  }

  // Create line chart - only use first 4 columns for chart
  const chartRange = chartSheet.getRange(startRow, 1, data.length, 4);
  const chart = chartSheet.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(chartRange)
    .setPosition(startRow, 8, 0, 0)
    .setOption("title", "Most Accurate RSVP Predictions (Top 15 Events)")
    .setOption("vAxis", { title: "Count" })
    .setOption("hAxis", { title: "Event Name", slantedText: true, slantedTextAngle: 45 })
    .setOption("series", {
      0: { color: "#4285F4", lineWidth: 2, labelInLegend: "RSVP" },
      1: { color: "#34A853", lineWidth: 2, labelInLegend: "Attended" },
      2: { color: "#EA4335", lineWidth: 1, lineDashStyle: [4, 4], labelInLegend: "Difference" }
    })
    .setOption("legend", { position: "bottom" })
    .setOption("curveType", "function")
    .build();
  chartSheet.insertChart(chart);

  return startRow + data.length + 2;
}

/**
 * Graph 3: All Events Timeline sorted by date
 * Line graph showing all events with RSVPs in chronological order
 */
function createAllEventsTimelineChart(chartSheet, startRow) {
  const events = extractEventData();

  // Filter events with RSVPs and sort by date
  const sortedEvents = events
    .filter(e => e.rsvp > 0) // Only include events with RSVPs
    .sort((a, b) => {
      // Sort by date - handle both Date objects and strings
      const dateA = a.date instanceof Date ? a.date : new Date(a.date);
      const dateB = b.date instanceof Date ? b.date : new Date(b.date);
      return dateA - dateB; // Ascending order (oldest first)
    });

  if (sortedEvents.length === 0) {
    chartSheet.getRange(startRow, 1).setValue("No events with RSVPs found");
    return startRow + 2;
  }

  // Write data to sheet - include date and day type
  const headers = ["Event Name", "RSVP", "Attended", "Difference", "Event Date", "Day Type"];
  const dataRows = sortedEvents.map(e => [
    e.name,
    e.rsvp,
    e.attended,
    e.difference,
    e.date,
    e.dayType
  ]);
  const data = [headers, ...dataRows];

  chartSheet.getRange(startRow, 1, data.length, headers.length).setValues(data);

  // Format numeric columns as numbers (not dates)
  chartSheet.getRange(startRow + 1, 2, dataRows.length, 3).setNumberFormat("0"); // RSVP, Attended, Difference
  // Format date column if they are Date objects
  if (sortedEvents[0].date instanceof Date) {
    chartSheet.getRange(startRow + 1, 5, dataRows.length, 1).setNumberFormat("mm/dd/yyyy");
  }

  // Create line chart - only use first 4 columns for chart
  const chartRange = chartSheet.getRange(startRow, 1, data.length, 4);
  const chart = chartSheet.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(chartRange)
    .setPosition(startRow, 8, 0, 0)
    .setOption("title", "All Events Timeline (Sorted by Date)")
    .setOption("vAxis", { title: "Count" })
    .setOption("hAxis", { title: "Event Name", slantedText: true, slantedTextAngle: 45 })
    .setOption("series", {
      0: { color: "#4285F4", lineWidth: 2, labelInLegend: "RSVP" },
      1: { color: "#34A853", lineWidth: 2, labelInLegend: "Attended" },
      2: { color: "#EA4335", lineWidth: 1, lineDashStyle: [4, 4], labelInLegend: "Difference" }
    })
    .setOption("legend", { position: "bottom" })
    .setOption("curveType", "function")
    .build();
  chartSheet.insertChart(chart);

  return startRow + data.length + 2;
}

/**
 * Graph 4: Overall RSVP vs Attended comparison
 * Excludes events with zero RSVPs from average calculations
 */
function createRSVPvsAttendedComparisonChart(chartSheet, startRow) {
  const events = extractEventData();

  // Filter out events with zero RSVPs
  const eventsWithRSVP = events.filter(e => e.rsvp > 0);

  if (eventsWithRSVP.length === 0) {
    chartSheet.getRange(startRow, 1).setValue("No events with RSVPs found");
    return startRow + 2;
  }

  // Calculate statistics
  const totalRSVP = eventsWithRSVP.reduce((sum, e) => sum + e.rsvp, 0);
  const totalAttended = eventsWithRSVP.reduce((sum, e) => sum + e.attended, 0);
  const avgRSVP = totalRSVP / eventsWithRSVP.length;
  const avgAttended = totalAttended / eventsWithRSVP.length;
  const attendanceRate = totalRSVP > 0 ? (totalAttended / totalRSVP * 100) : 0;

  // Write summary data
  const headers = ["Metric", "RSVP", "Attended"];
  const data = [
    headers,
    ["Total", totalRSVP, totalAttended],
    ["Average per Event", avgRSVP, avgAttended],
    ["Event Count", eventsWithRSVP.length, eventsWithRSVP.length],
    ["Attendance Rate (%)", "", attendanceRate]
  ];

  chartSheet.getRange(startRow, 1, data.length, headers.length).setValues(data);

  // Format numeric columns as numbers (not dates)
  chartSheet.getRange(startRow + 1, 2, 3, 2).setNumberFormat("0.00"); // RSVP and Attended for Total, Average, Count rows
  chartSheet.getRange(startRow + 4, 3, 1, 1).setNumberFormat("0.00"); // Attendance rate percentage

  // Create comparison chart (using Total and Average rows)
  const chartRange = chartSheet.getRange(startRow, 1, 3, 3); // Headers + Total + Average
  const chart = chartSheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(chartRange)
    .setPosition(startRow, 5, 0, 0)
    .setOption("title", "RSVP vs Actual Attendance Comparison\n(Events with RSVPs only)")
    .setOption("vAxis", { title: "Count" })
    .setOption("hAxis", { title: "Metric" })
    .setOption("series", {
      0: { color: "#4285F4", labelInLegend: "RSVP" },
      1: { color: "#34A853", labelInLegend: "Attended" }
    })
    .setOption("legend", { position: "bottom" })
    .setOption("isStacked", false)
    .build();
  chartSheet.insertChart(chart);

  // Write detailed event breakdown below (optional columns)
  const breakdownStartRow = startRow + data.length + 2;
  chartSheet.getRange(breakdownStartRow, 1).setValue("Detailed Event Breakdown:");
  chartSheet.getRange(breakdownStartRow, 1).setFontWeight("bold");

  const breakdownHeaders = ["Event Name", "RSVP", "Attended", "Event Date", "Day Type"];
  const breakdownRows = eventsWithRSVP.map(e => [e.name, e.rsvp, e.attended, e.date, e.dayType]);
  const breakdownData = [breakdownHeaders, ...breakdownRows];

  chartSheet.getRange(breakdownStartRow + 1, 1, breakdownData.length, breakdownHeaders.length).setValues(breakdownData);
  chartSheet.getRange(breakdownStartRow + 2, 2, breakdownRows.length, 2).setNumberFormat("0"); // RSVP and Attended

  // Format date column if they are Date objects
  if (eventsWithRSVP.length > 0 && eventsWithRSVP[0].date instanceof Date) {
    chartSheet.getRange(breakdownStartRow + 2, 4, breakdownRows.length, 1).setNumberFormat("mm/dd/yyyy");
  }

  return breakdownStartRow + breakdownData.length + 2;
}

/**
 * Main function to create all four analysis charts
 */
function createRSVPvsAttendanceChart() {
  const [contactListSheet] = sheetsByName();
  const ss = contactListSheet.getParent();
  const chartSheetName = "Data Analysis Graphs";

  // Get or create chart sheet
  let chartSheet = ss.getSheetByName(chartSheetName);
  if (!chartSheet) {
    chartSheet = ss.insertSheet(chartSheetName);
  }

  // Clear previous content
  chartSheet.clear();

  // Add title
  chartSheet.getRange(1, 1).setValue("RSVP & Attendance Analysis Dashboard");
  chartSheet.getRange(1, 1).setFontSize(14).setFontWeight("bold");

  let currentRow = 3;

  // Create Graph 1: Average RSVP by Day
  currentRow = createAverageRSVPByDayChart(chartSheet, currentRow);

  // Create Graph 2: Most Accurate Predictions (Top 15)
  currentRow = createRSVPAccuracyChart(chartSheet, currentRow);

  // Create Graph 3: All Events Timeline (Sorted by Date)
  currentRow = createAllEventsTimelineChart(chartSheet, currentRow);

  // Create Graph 4: Overall Comparison
  currentRow = createRSVPvsAttendedComparisonChart(chartSheet, currentRow);

  Logger.log("All four charts created successfully in 'Data Analysis Graphs' sheet");
}