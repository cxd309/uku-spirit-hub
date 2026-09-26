/**
 * name of the Club Report tab
 */
const CLUB_REPORT_SHEET = "Club Report";

/**
 * columns on the Club Report tab
 */
const CLUB_REPORT_WIDTH = 14;

/**
 * text for the info row at the top of the Club Report tab
 */
const CLUB_REPORT_INFO = "The spirit scores one club received\n\n"
  + "Pick a club in the Club cell, everything below updates straight away\n"
  + "Counts the same scores as the award: included tournaments, between different clubs\n"
  + "Mean and Rank are from the Club Statistics tab, as last calculated";

/**
 * where each part of the Club Report tab sits, 1-based rows
 */
const CLUB_REPORT_ROWS = Object.freeze({
  club: 3,
  summary: 4,
  totals: 11,
  categories: 35,
});

/**
 * the summary lines under the club, label then what they look up
 */
const CLUB_REPORT_SUMMARY = Object.freeze(["Teams", "Tournaments", "Scores counted", "Mean", "Rank"]);

/**
 * colours for scores 0 to 4 in the category chart
 * diverging: reds for low scores, grey for the middle, blues for high scores
 * each arm checked for monotone lightness and at least 2:1 contrast on white
 */
const CLUB_REPORT_SCORE_COLOURS = Object.freeze(["#b83232", "#eb8a86", "#b5b4b0", "#86b6ef", "#1c5cab"]);

/**
 * colour of the single series in the totals chart
 */
const CLUB_REPORT_TOTAL_COLOUR = "#2a78d6";

/**
 * formula counting the club's counted scores that match one condition
 *
 * @param {keyof typeof RESPONSE_HEADERS} column  Results column to test
 * @param {string}                        value   cell holding the value to match
 * @returns {string} the formula
 */
function _clubReportCount_(column, value) {
  /** @param {keyof typeof RESPONSE_HEADERS} key */
  const results = (key) => _columnBelowHeader_(RESPONSES_SHEET, RESPONSE_KEYS.indexOf(key) + 1);
  const club = `$B$${CLUB_REPORT_ROWS.club}`;
  return `=COUNTIFS(${results("receiverClub")}, ${club}, ${results("countsForAward")}, TRUE, ${
    results(column)
  }, ${value})`;
}

/**
 * summary formulas, one per CLUB_REPORT_SUMMARY line
 *
 * @returns {string[]} the formulas
 */
function _clubReportSummaryFormulas_() {
  const club = `$B$${CLUB_REPORT_ROWS.club}`;
  /** @param {keyof typeof CLUB_HEADERS} key */
  const clubs = (key) => _columnBelowHeader_(CLUBS_SHEET, CLUB_KEYS.indexOf(key) + 1);
  /** @param {keyof typeof CLUB_STATS_HEADERS} key */
  const stats = (key) => _columnBelowHeader_(CLUB_STATS_SHEET, CLUB_STATS_KEYS.indexOf(key) + 1);
  /** @param {keyof typeof RESPONSE_HEADERS} key */
  const results = (key) => _columnBelowHeader_(RESPONSES_SHEET, RESPONSE_KEYS.indexOf(key) + 1);
  return [
    `=IFERROR(XLOOKUP(${club}, ${clubs("club")}, ${clubs("teams")}), "")`,
    `=IFERROR(XLOOKUP(${club}, ${clubs("club")}, ${clubs("events")}), "")`,
    `=COUNTIFS(${results("receiverClub")}, ${club}, ${results("countsForAward")}, TRUE)`,
    `=IFERROR(XLOOKUP(${club}, ${stats("club")}, ${stats("mean")}), "")`,
    `=IFERROR(XLOOKUP(${club}, ${stats("club")}, ${stats("rank")}), "")`,
  ];
}

/**
 * build the Club Report tab: the club picker, summary, two tables of counts and their charts
 * everything is formulas, so picking another club needs no script
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet  a new, empty tab
 */
function _buildClubReport_(sheet) {
  const rows = CLUB_REPORT_ROWS;
  if (sheet.getMaxColumns() > CLUB_REPORT_WIDTH) {
    sheet.deleteColumns(CLUB_REPORT_WIDTH + 1, sheet.getMaxColumns() - CLUB_REPORT_WIDTH);
  }
  sheet
    .getRange(INFO_ROW, 1, 1, CLUB_REPORT_WIDTH)
    .merge()
    .setValue(CLUB_REPORT_INFO)
    .setWrap(true)
    .setVerticalAlignment("top");

  sheet.getRange(rows.club, 1, 1, 2).setValues([["Club", ""]]).setFontWeight("bold");
  sheet.getRange(rows.club, 2, 1, 5).merge();
  const formulas = _clubReportSummaryFormulas_();
  sheet
    .getRange(rows.summary, 1, CLUB_REPORT_SUMMARY.length, 2)
    .setValues(CLUB_REPORT_SUMMARY.map((label, i) => [label, formulas[i]]));
  sheet.getRange(rows.summary, 1, CLUB_REPORT_SUMMARY.length, 1).setFontWeight("bold");
  for (let i = 0; i < CLUB_REPORT_SUMMARY.length; i++) {
    sheet.getRange(rows.summary + i, 2, 1, 6).merge().setWrap(true).setHorizontalAlignment("left");
  }
  sheet.getRange(rows.summary + CLUB_REPORT_SUMMARY.indexOf("Mean"), 2).setNumberFormat("0.00");

  // totals received, 0 to 20, labels as text so the chart treats them as categories
  const totals = Array.from({ length: 21 }, (_, t) => t);
  sheet.getRange(rows.totals, 1, 1, 2).setValues([["Total", "Times received"]]).setFontWeight("bold");
  sheet.getRange(rows.totals + 1, 1, totals.length, 1).setNumberFormat("@");
  sheet.getRange(rows.totals + 1, 1, totals.length, 2).setValues(
    totals.map((t, i) => [String(t), _clubReportCount_("total", `$A${rows.totals + 1 + i}`)]),
  );

  // each category, scores 0 to 4 across
  const scores = [0, 1, 2, 3, 4];
  sheet
    .getRange(rows.categories, 1, 1, 1 + scores.length)
    .setValues([["Category", ...scores.map((s) => `Score ${s}`)]])
    .setFontWeight("bold");
  sheet.getRange(rows.categories + 1, 1, SCORE_KEYS.length, 1 + scores.length).setValues(
    SCORE_KEYS.map((key) => [RESPONSE_HEADERS[key], ...scores.map((s) => _clubReportCount_(key, String(s)))]),
  );

  const totalsChart = sheet
    .newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(sheet.getRange(rows.totals, 1, totals.length + 1, 2))
    .setNumHeaders(1)
    .setPosition(rows.totals, 4, 0, 0)
    .setOption("title", "Total scores received")
    .setOption("legend", { position: "none" })
    .setOption("colors", [CLUB_REPORT_TOTAL_COLOUR])
    .setOption("hAxis", { title: "Total score (0 to 20)" })
    .setOption("vAxis", { title: "Times received", minValue: 0, format: "0" })
    .setOption("width", 720)
    .setOption("height", 420)
    .build();
  sheet.insertChart(totalsChart);

  const categoriesChart = sheet
    .newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(sheet.getRange(rows.categories, 1, SCORE_KEYS.length + 1, 1 + scores.length))
    .setNumHeaders(1)
    .setPosition(rows.categories, 8, 0, 0)
    .setOption("title", "Category scores received")
    .setOption("isStacked", true)
    .setOption("legend", { position: "right" })
    .setOption("colors", [...CLUB_REPORT_SCORE_COLOURS])
    .setOption("vAxis", { title: "Times received", minValue: 0, format: "0" })
    .setOption("width", 620)
    .setOption("height", 420)
    .build();
  sheet.insertChart(categoriesChart);
}

/**
 * get the Club Report tab, building it on first use
 * an existing tab is left exactly as it is, so a club picked by a person stays picked
 *
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} the Club Report tab
 */
function _getClubReportSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const existing = ss.getSheetByName(CLUB_REPORT_SHEET);
  if (existing) return existing;
  const sheet = ss.insertSheet(CLUB_REPORT_SHEET);
  _buildClubReport_(sheet);
  _setClubReportChoices_(_clubNames_([..._readTeamClubs_().values()]));
  return sheet;
}

/**
 * make the Club cell a dropdown of every club
 * called whenever Clubs is rebuilt, does nothing until the tab exists
 *
 * @param {string[]} clubs  club names, in display order
 */
function _setClubReportChoices_(clubs) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CLUB_REPORT_SHEET);
  if (!sheet) return;
  const cell = sheet.getRange(CLUB_REPORT_ROWS.club, 2);
  if (clubs.length === 0) {
    cell.clearDataValidations();
    return;
  }
  cell.setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(clubs, true).setAllowInvalid(false).build(),
  );
}
