/**
 * name of the Club Report tab
 */
const CLUB_REPORT_SHEET = "Club Report";

/**
 * most tournaments in a totals chart, and most teams in the team picker, later ones are left off
 */
const CLUB_REPORT_SERIES = 40;

/**
 * columns on the Club Report tab, the tables below the charts need a label column and one per tournament
 */
const CLUB_REPORT_WIDTH = 1 + CLUB_REPORT_SERIES;

/**
 * columns the info row is merged across, the width of the charts
 */
const CLUB_REPORT_INFO_WIDTH = 14;

/**
 * text for the info row at the top of the Club Report tab
 */
const CLUB_REPORT_INFO = "The spirit scores one club, and one of its teams, received\n\n"
  + "Pick a club in the Club cell and a team in the Team cell, everything below updates straight away\n"
  + "Counts the same scores as the award: see Counted on the Club Statistics tab\n"
  + "Teams are matched without trailing brackets, so Birmingham 1 (WW4) and Birmingham 1 (WW2) are both Birmingham 1\n"
  + "Mean and Rank are from the Club Statistics tab, as last calculated\n"
  + "The charts are drawn from the tables below them";

/**
 * where each part of the Club Report tab sits, 1-based rows
 */
const CLUB_REPORT_ROWS = Object.freeze({
  club: 3,
  summary: 4,
  clubCharts: 10,
  team: 31,
  teamSummary: 32,
  teamCharts: 36,
  clubCategories: 58,
  teamCategories: 65,
  clubTotals: 72,
  teams: 96,
  teamTotals: 98,
});

/**
 * rows of a totals table, counted from its first row
 * the date and tournament rows are one formula spilling across, the tournament row is also the chart's header
 */
const CLUB_REPORT_TOTALS_TABLE = Object.freeze({ date: 0, header: 1, totals: 2 });

/**
 * totals a score can have, 0 to 20
 */
const CLUB_REPORT_TOTALS = 21;

/**
 * chart titles, the totals titles are also how those charts are found again to fit their range
 */
const CLUB_REPORT_TITLES = Object.freeze({
  clubTotals: "Club: total scores received, by tournament",
  clubCategories: "Club: category scores received",
  teamTotals: "Team: total scores received, by tournament",
  teamCategories: "Team: category scores received",
});

/**
 * the summary lines under the club, label then what they look up
 */
const CLUB_REPORT_SUMMARY = Object.freeze(["Teams", "Tournaments", "Scores counted", "Mean", "Rank"]);

/**
 * the summary lines under the team
 */
const CLUB_REPORT_TEAM_SUMMARY = Object.freeze(["Tournaments", "Scores counted", "Mean"]);

/**
 * colours for scores 0 to 4 in the category charts
 * diverging: reds for low scores, grey for the middle, blues for high scores
 * each arm checked for monotone lightness and at least 2:1 contrast on white
 */
const CLUB_REPORT_SCORE_COLOURS = Object.freeze(["#b83232", "#eb8a86", "#b5b4b0", "#86b6ef", "#1c5cab"]);

/**
 * a Results column below its header, for formulas
 *
 * @param key  the column
 * @returns the reference
 */
function _clubReportResults_(key: keyof typeof RESPONSE_HEADERS): string {
  return _columnBelowHeader_(RESPONSES_SHEET, RESPONSE_KEYS.indexOf(key) + 1);
}

/**
 * the conditions picking the scores a table counts
 * the club's counted scores, and only the picked team's when for the team
 *
 * @param forTeam  true for the team tables
 * @returns Results column and the cell or value it must match
 */
function _clubReportConditions_(forTeam: boolean): [keyof typeof RESPONSE_HEADERS, string][] {
  const conditions: [keyof typeof RESPONSE_HEADERS, string][] = [
    ["receiverClub", `$B$${CLUB_REPORT_ROWS.club}`],
    ["countsForAward", "TRUE"],
  ];
  return forTeam ? [...conditions, ["reportTeam", `$B$${CLUB_REPORT_ROWS.team}`]] : conditions;
}

/**
 * the conditions as COUNTIFS arguments
 *
 * @param conditions  from _clubReportConditions_
 * @returns e.g. 'Results'!$J$3:$J, $B$3, ...
 */
function _clubReportCriteria_(conditions: [keyof typeof RESPONSE_HEADERS, string][]): string {
  return conditions.map(([key, value]) => `${_clubReportResults_(key)}, ${value}`).join(", ");
}

/**
 * the conditions as FILTER arguments
 *
 * @param conditions  from _clubReportConditions_
 * @returns e.g. 'Results'!$J$3:$J=$B$3, ...
 */
function _clubReportFilter_(conditions: [keyof typeof RESPONSE_HEADERS, string][]): string {
  return conditions.map(([key, value]) => `${_clubReportResults_(key)}=${value}`).join(", ");
}

/**
 * summary formulas, one per CLUB_REPORT_SUMMARY line
 *
 * @returns the formulas
 */
function _clubReportSummaryFormulas_(): string[] {
  const club = `$B$${CLUB_REPORT_ROWS.club}`;
  const clubs = (key: keyof typeof CLUB_HEADERS) => _columnBelowHeader_(CLUBS_SHEET, CLUB_KEYS.indexOf(key) + 1);
  const stats = (key: keyof typeof CLUB_STATS_HEADERS) =>
    _columnBelowHeader_(CLUB_STATS_SHEET, CLUB_STATS_KEYS.indexOf(key) + 1);
  return [
    `=IFERROR(XLOOKUP(${club}, ${clubs("club")}, ${clubs("teams")}), "")`,
    `=IFERROR(XLOOKUP(${club}, ${clubs("club")}, ${clubs("events")}), "")`,
    `=COUNTIFS(${_clubReportCriteria_(_clubReportConditions_(false))})`,
    `=IFERROR(XLOOKUP(${club}, ${stats("club")}, ${stats("mean")}), "")`,
    `=IFERROR(XLOOKUP(${club}, ${stats("club")}, ${stats("rank")}), "")`,
  ];
}

/**
 * team summary formulas, one per CLUB_REPORT_TEAM_SUMMARY line
 *
 * @returns the formulas
 */
function _clubReportTeamSummaryFormulas_(): string[] {
  const criteria = _clubReportCriteria_(_clubReportConditions_(true));
  const header = CLUB_REPORT_ROWS.teamTotals + CLUB_REPORT_TOTALS_TABLE.header;
  return [
    `=TEXTJOIN(", ", TRUE, B${header}:${_columnLetter_(CLUB_REPORT_WIDTH)}${header})`,
    `=COUNTIFS(${criteria})`,
    `=IFERROR(AVERAGEIFS(${_clubReportResults_("total")}, ${criteria}), "")`,
  ];
}

/**
 * labels in column A and formulas in column B, each line merged across so long text wraps
 *
 * @param sheet     the Club Report tab
 * @param row       first line
 * @param labels    one per line
 * @param formulas  one per line
 */
function _clubReportLines_(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  row: number,
  labels: readonly string[],
  formulas: string[],
) {
  sheet.getRange(row, 1, labels.length, 2).setValues(labels.map((label, i) => [label, formulas[i]]));
  sheet.getRange(row, 1, labels.length, 1).setFontWeight("bold");
  for (let i = 0; i < labels.length; i++) {
    sheet.getRange(row + i, 2, 1, 6).merge().setWrap(true).setHorizontalAlignment("left");
  }
  const mean = labels.indexOf("Mean");
  if (mean >= 0) sheet.getRange(row + mean, 2).setNumberFormat("0.00");
}

/**
 * a categories table: each category, how many times each score 0 to 4 was received
 *
 * @param sheet    the Club Report tab
 * @param top      the table's header row
 * @param forTeam  true for the team's table
 * @returns the table, header included, for the chart
 */
function _clubReportCategoriesTable_(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  top: number,
  forTeam: boolean,
): GoogleAppsScript.Spreadsheet.Range {
  const scores = [0, 1, 2, 3, 4];
  const criteria = _clubReportCriteria_(_clubReportConditions_(forTeam));
  sheet
    .getRange(top, 1, 1, 1 + scores.length)
    .setValues([[forTeam ? "Team category" : "Club category", ...scores.map((s) => `Score ${s}`)]])
    .setFontWeight("bold");
  sheet.getRange(top + 1, 1, SCORE_KEYS.length, 1 + scores.length).setValues(
    SCORE_KEYS.map((key) => [
      RESPONSE_HEADERS[key],
      ...scores.map((s) => `=COUNTIFS(${criteria}, ${_clubReportResults_(key)}, ${s})`),
    ]),
  );
  return sheet.getRange(top, 1, SCORE_KEYS.length + 1, 1 + scores.length);
}

/**
 * a totals table: one column per tournament with counted scores, in date order
 * then how many times each total 0 to 20 was received there
 * all formulas, so picking another club or team needs no script
 *
 * @param sheet    the Club Report tab
 * @param top      the table's first row
 * @param forTeam  true for the team's table
 */
function _clubReportTotalsTable_(sheet: GoogleAppsScript.Spreadsheet.Sheet, top: number, forTeam: boolean) {
  const at = CLUB_REPORT_TOTALS_TABLE;
  const conditions = _clubReportConditions_(forTeam);
  const header = top + at.header;
  const first = top + at.totals;

  sheet.getRange(top, 1, 2, 1).setValues([["Date"], ["Total"]]).setFontWeight("bold");
  sheet.getRange(top + at.date, 2).setFormula(
    `=IFERROR(ARRAY_CONSTRAIN(TRANSPOSE(SORT(UNIQUE(FILTER({${_clubReportResults_("date")}, ${
      _clubReportResults_("tournament")
    }}, ${_clubReportFilter_(conditions)})), 1, TRUE, 2, TRUE)), 2, ${CLUB_REPORT_SERIES}), "")`,
  );
  sheet.getRange(top + at.date, 2, 1, CLUB_REPORT_SERIES).setNumberFormat("yyyy-mm-dd");
  sheet.getRange(header, 2, 1, CLUB_REPORT_SERIES).setFontWeight("bold");

  // labels as text so the chart treats the totals as categories
  sheet.getRange(first, 1, CLUB_REPORT_TOTALS, 1).setNumberFormat("@").setValues(
    Array.from({ length: CLUB_REPORT_TOTALS }, (_, t) => [String(t)]),
  );
  const criteria = _clubReportCriteria_(conditions);
  const tournament = _clubReportResults_("tournament");
  const total = _clubReportResults_("total");
  const columns = Array.from({ length: CLUB_REPORT_SERIES }, (_, i) => _columnLetter_(i + 2));
  sheet.getRange(first, 2, CLUB_REPORT_TOTALS, CLUB_REPORT_SERIES).setFormulas(
    Array.from(
      { length: CLUB_REPORT_TOTALS },
      (_, t) =>
        columns.map((c) =>
          `=IF(${c}$${header}="", "", COUNTIFS(${criteria}, ${tournament}, ${c}$${header}, ${total}, $A${first + t}))`
        ),
    ),
  );
}

/**
 * a totals table's tournament and count rows, for its chart
 *
 * @param sheet   the Club Report tab
 * @param top     the table's first row
 * @param series  how many tournament columns to cover
 * @returns the range, label column included
 */
function _clubReportTotalsRange_(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  top: number,
  series: number,
): GoogleAppsScript.Spreadsheet.Range {
  return sheet.getRange(top + CLUB_REPORT_TOTALS_TABLE.header, 1, 1 + CLUB_REPORT_TOTALS, 1 + series);
}

/**
 * add a totals chart and a categories chart side by side
 *
 * @param sheet       the Club Report tab
 * @param row         row the charts sit on
 * @param totals      first row of the totals table
 * @param categories  the categories table
 * @param forTeam     true for the team's charts
 */
function _clubReportCharts_(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  row: number,
  totals: number,
  categories: GoogleAppsScript.Spreadsheet.Range,
  forTeam: boolean,
) {
  sheet.insertChart(
    sheet
      .newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(_clubReportTotalsRange_(sheet, totals, 1))
      .setNumHeaders(1)
      .setPosition(row, 1, 0, 0)
      .setOption("title", forTeam ? CLUB_REPORT_TITLES.teamTotals : CLUB_REPORT_TITLES.clubTotals)
      .setOption("isStacked", true)
      .setOption("legend", { position: "right" })
      .setOption("hAxis", { title: "Total score (0 to 20)" })
      .setOption("vAxis", { title: "Times received", minValue: 0, format: "0" })
      .setOption("width", 680)
      .setOption("height", 400)
      .build(),
  );
  sheet.insertChart(
    sheet
      .newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(categories)
      .setNumHeaders(1)
      .setPosition(row, 8, 0, 0)
      .setOption("title", forTeam ? CLUB_REPORT_TITLES.teamCategories : CLUB_REPORT_TITLES.clubCategories)
      .setOption("isStacked", true)
      .setOption("legend", { position: "right" })
      .setOption("colors", [...CLUB_REPORT_SCORE_COLOURS])
      .setOption("vAxis", { title: "Times received", minValue: 0, format: "0" })
      .setOption("width", 560)
      .setOption("height", 400)
      .build(),
  );
}

/**
 * build the Club Report tab: the club and team pickers, their summaries, tables and charts
 * everything is formulas, so picking another club or team needs no script, only the totals charts are refitted
 *
 * @param sheet  a new, empty tab
 */
function _buildClubReport_(sheet: GoogleAppsScript.Spreadsheet.Sheet) {
  const rows = CLUB_REPORT_ROWS;
  _fitSheet_(sheet, rows.teamTotals + CLUB_REPORT_TOTALS_TABLE.totals + CLUB_REPORT_TOTALS - 1, CLUB_REPORT_WIDTH);
  sheet
    .getRange(INFO_ROW, 1, 1, CLUB_REPORT_INFO_WIDTH)
    .merge()
    .setValue(CLUB_REPORT_INFO)
    .setWrap(true)
    .setVerticalAlignment("top");

  for (const [row, label] of [[rows.club, "Club"], [rows.team, "Team"]] as const) {
    sheet.getRange(row, 1, 1, 2).setValues([[label, ""]]).setFontWeight("bold");
    sheet.getRange(row, 2, 1, 5).merge();
  }
  _clubReportLines_(sheet, rows.summary, CLUB_REPORT_SUMMARY, _clubReportSummaryFormulas_());
  _clubReportLines_(sheet, rows.teamSummary, CLUB_REPORT_TEAM_SUMMARY, _clubReportTeamSummaryFormulas_());

  // the club's teams, without trailing brackets, for the Team picker
  sheet.getRange(rows.teams, 1).setValue("Teams").setFontWeight("bold");
  sheet.getRange(rows.teams, 2).setFormula(
    `=IFERROR(ARRAY_CONSTRAIN(TRANSPOSE(SORT(UNIQUE(FILTER(${_clubReportResults_("reportTeam")}, ${
      _clubReportFilter_(_clubReportConditions_(false))
    })))), 1, ${CLUB_REPORT_SERIES}), "")`,
  );
  sheet.getRange(rows.team, 2).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInRange(sheet.getRange(rows.teams, 2, 1, CLUB_REPORT_SERIES), true)
      .setAllowInvalid(true)
      .build(),
  );

  const clubCategories = _clubReportCategoriesTable_(sheet, rows.clubCategories, false);
  const teamCategories = _clubReportCategoriesTable_(sheet, rows.teamCategories, true);
  _clubReportTotalsTable_(sheet, rows.clubTotals, false);
  _clubReportTotalsTable_(sheet, rows.teamTotals, true);
  _clubReportCharts_(sheet, rows.clubCharts, rows.clubTotals, clubCategories, false);
  _clubReportCharts_(sheet, rows.teamCharts, rows.teamTotals, teamCategories, true);
}

/**
 * fit each totals chart to its tournament columns
 * a chart shows every column it covers, even empty ones, so its range has to change with the club and team
 * waits for the tables' formulas, so only call when the club, team or scores may have changed
 * does nothing until the tab exists
 *
 * @param pickFirstTeam  also pick the club's first team, for when the club has changed
 */
function _fitClubReportCharts_(pickFirstTeam = false) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CLUB_REPORT_SHEET);
  if (!sheet) return;
  const rows = CLUB_REPORT_ROWS;
  SpreadsheetApp.flush();
  if (pickFirstTeam) {
    sheet.getRange(rows.team, 2).setValue(sheet.getRange(rows.teams, 2).getValue());
    SpreadsheetApp.flush();
  }
  const charts = sheet.getCharts();
  for (
    const [title, top] of [[CLUB_REPORT_TITLES.clubTotals, rows.clubTotals], [
      CLUB_REPORT_TITLES.teamTotals,
      rows.teamTotals,
    ]] as const
  ) {
    const chart = charts.find((c) => c.getOptions().get("title") === title);
    if (!chart) continue;
    const headers = sheet.getRange(top + CLUB_REPORT_TOTALS_TABLE.header, 2, 1, CLUB_REPORT_SERIES).getValues()[0];
    const series = Math.max(1, headers.filter((h) => h !== "").length);
    sheet.updateChart(chart.modify().clearRanges().addRange(_clubReportTotalsRange_(sheet, top, series)).build());
  }
}

/**
 * get the Club Report tab, building it on first use
 * an existing tab is left exactly as it is, so a club picked by a person stays picked
 * a tab from before the team section (0.12) is built again in the same place, keeping the picked club
 *
 * @returns the Club Report tab
 */
function _getClubReportSheet_(): GoogleAppsScript.Spreadsheet.Sheet {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const existing = ss.getSheetByName(CLUB_REPORT_SHEET);
  const current = existing !== null
    && existing.getMaxRows() >= CLUB_REPORT_ROWS.team
    && existing.getRange(CLUB_REPORT_ROWS.team, 1).getValue() === "Team";
  if (existing && current) return existing;

  const club = existing ? existing.getRange(CLUB_REPORT_ROWS.club, 2).getValue() : "";
  const index = existing ? existing.getIndex() - 1 : ss.getNumSheets();
  if (existing) ss.deleteSheet(existing);
  const sheet = ss.insertSheet(CLUB_REPORT_SHEET, index);
  _buildClubReport_(sheet);
  _setClubReportChoices_(_clubNames_([..._readTeamClubs_().values()]));
  if (club !== "") sheet.getRange(CLUB_REPORT_ROWS.club, 2).setValue(club);
  _fitClubReportCharts_(club !== "");
  return sheet;
}

/**
 * make the Club cell a dropdown of every club
 * called whenever Clubs is rebuilt, does nothing until the tab exists
 *
 * @param clubs  club names, in display order
 */
function _setClubReportChoices_(clubs: string[]) {
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
