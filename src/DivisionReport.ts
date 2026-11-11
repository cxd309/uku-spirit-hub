/**
 * name of the Division Report tab
 */
const DIVISION_REPORT_SHEET = "Division Report";

/**
 * the chart groups on the Division Report tab, top to bottom
 * All is every tournament, the others only tournaments with that Div on the Tournaments tab
 * a function so EVENT_DIVISIONS has loaded before it is read
 *
 * @returns the group labels
 */
function _divisionReportGroups_(): string[] {
  return ["All", ...EVENT_DIVISIONS];
}

/**
 * rows each group's heading and charts take, the charts are 400px tall
 */
const DIVISION_REPORT_GROUP_ROWS = 23;

/**
 * rows each group's tables take: categories header and rows, a gap, totals date, header and rows, a gap
 * a function so SCORE_KEYS has loaded before it is read
 *
 * @returns the rows
 */
function _divisionReportTableRows_(): number {
  return 1 + SCORE_KEYS.length + 1 + 2 + CLUB_REPORT_TOTALS + 1;
}

/**
 * the first row of the charts, below the info row and a gap
 */
const DIVISION_REPORT_FIRST_ROW = 3;

/**
 * info row text, with when it was calculated
 *
 * @param when  when the report was calculated
 * @returns the info text
 */
function _divisionReportInfo_(when: Date): string {
  return `The spirit scores received at each division's tournaments, calculated ${
    Utilities.formatDate(when, Session.getScriptTimeZone(), "d MMM yyyy HH:mm")
  }\n\n`
    + `Set each tournament's Div (${EVENT_DIVISIONS.join(", ")}) on the Tournaments tab, `
    + "tournaments with no Div are only in All\n"
    + "Counts the same scores as the award: see Counted on the Club Statistics tab\n"
    + "The charts are drawn from the tables below them\n\n"
    + "DO NOT EDIT, run \"Refresh Club Statistics\" to recalculate";
}

/**
 * one chart group of the Division Report: its tournaments, and the scores received at them
 */
interface DivisionGroup {
  /** All, or a Div */
  label: string;
  /** tournaments with counted scores, in date order */
  tournaments: { name: string; date: Date | null }[];
  /** [total 0 to 20][tournament] → times received */
  totals: number[][];
  /** [category, in SCORE_KEYS order][score 0 to 4] → times received */
  categories: number[][];
  /** scores counted */
  count: number;
  /** mean total, "" with no scores */
  mean: number | "";
}

/**
 * every chart group of the Division Report
 * pure, never modifies its arguments
 *
 * @param events  all events, in date order
 * @param scores  every score at an included tournament, only those that count are used
 * @returns one group per _divisionReportGroups_
 */
function _divisionGroups_(events: EventRecord[], scores: CountedScore[]): DivisionGroup[] {
  const counted = scores.filter((s) => s.counts);
  return _divisionReportGroups_().map((label) => {
    const inGroup = events.filter((e) => label === "All" || e.division === label);
    const byTournament = inGroup
      .map((event) => ({ event, scores: counted.filter((s) => s.tournament === event.spiritFormUrl) }))
      .filter((t) => t.scores.length > 0);
    const all = byTournament.flatMap((t) => t.scores);
    return {
      label,
      tournaments: byTournament.map((t) => ({ name: t.event.eventName, date: t.event.startDate })),
      totals: Array.from(
        { length: CLUB_REPORT_TOTALS },
        (_, total) => byTournament.map((t) => t.scores.filter((s) => s.total === total).length),
      ),
      categories: SCORE_KEYS.map((_, c) =>
        [0, 1, 2, 3, 4].map((score) => all.filter((s) => s.categories[c] === score).length)
      ),
      count: all.length,
      mean: _summarise_(all.map((s) => s.total)).mean,
    };
  });
}

/**
 * the line next to a group's label
 *
 * @param group  the group
 * @returns e.g. "12 tournaments, 845 scores counted, mean 11.24"
 */
function _divisionSummary_(group: DivisionGroup): string {
  if (group.tournaments.length === 0) {
    return group.label === "All"
      ? "No counted scores yet"
      : `No tournament with Div ${group.label} has counted scores yet`;
  }
  const mean = typeof group.mean === "number" ? `, mean ${group.mean.toFixed(2)}` : "";
  return `${group.tournaments.length} tournament(s), ${group.count} score(s) counted${mean}`;
}

/**
 * @returns the Division Report tab, created on first use
 */
function _getDivisionReportSheet_(): GoogleAppsScript.Spreadsheet.Sheet {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const existing = ss.getSheetByName(DIVISION_REPORT_SHEET);
  if (existing) return existing;
  const after = ss.getSheetByName(CLUB_REPORT_SHEET);
  const sheet = ss.insertSheet(DIVISION_REPORT_SHEET, after ? after.getIndex() : ss.getNumSheets());
  _fitSheet_(sheet, DIVISION_REPORT_FIRST_ROW, CLUB_REPORT_INFO_WIDTH);
  sheet
    .getRange(INFO_ROW, 1, 1, CLUB_REPORT_INFO_WIDTH)
    .merge()
    .setValue(
      "The spirit scores received at each division's tournaments\n\nRun \"Refresh Club Statistics\" to calculate",
    )
    .setWrap(true)
    .setVerticalAlignment("top");
  return sheet;
}

/**
 * write one group's tables
 *
 * @param sheet  the Division Report tab
 * @param top    the group's first table row
 * @param group  the group
 * @returns the categories table and the totals table, headers included, for the charts
 */
function _writeDivisionTables_(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  top: number,
  group: DivisionGroup,
): { categories: GoogleAppsScript.Spreadsheet.Range; totals: GoogleAppsScript.Spreadsheet.Range } {
  const scores = [0, 1, 2, 3, 4];
  const categories = sheet.getRange(top, 1, 1 + SCORE_KEYS.length, 1 + scores.length);
  categories.setValues([
    [`${group.label} category`, ...scores.map((s) => `Score ${s}`)],
    ...SCORE_KEYS.map((key, c) => [RESPONSE_HEADERS[key], ...group.categories[c]]),
  ]);
  sheet.getRange(top, 1, 1, 1 + scores.length).setFontWeight("bold");

  const date = top + 1 + SCORE_KEYS.length + 1;
  const n = group.tournaments.length;
  sheet.getRange(date, 1, 2, 1).setValues([["Date"], ["Total"]]).setFontWeight("bold");
  // labels as text so the chart treats the totals as categories
  sheet.getRange(date + 2, 1, CLUB_REPORT_TOTALS, 1).setNumberFormat("@").setValues(
    Array.from({ length: CLUB_REPORT_TOTALS }, (_, t) => [String(t)]),
  );
  if (n > 0) {
    sheet
      .getRange(date, 2, 2, n)
      .setValues([group.tournaments.map((t) => t.date ?? ""), group.tournaments.map((t) => t.name)]);
    sheet.getRange(date, 2, 1, n).setNumberFormat("yyyy-mm-dd");
    sheet.getRange(date + 1, 2, 1, n).setFontWeight("bold");
    sheet.getRange(date + 2, 2, CLUB_REPORT_TOTALS, n).setValues(group.totals);
  }
  return { categories, totals: sheet.getRange(date + 1, 1, 1 + CLUB_REPORT_TOTALS, 1 + n) };
}

/**
 * a group's totals chart and categories chart side by side, as on the Club Report
 *
 * @param sheet   the Division Report tab
 * @param row     row the charts sit on
 * @param label   the group's label
 * @param tables  from _writeDivisionTables_
 */
function _insertDivisionCharts_(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  row: number,
  label: string,
  tables: { categories: GoogleAppsScript.Spreadsheet.Range; totals: GoogleAppsScript.Spreadsheet.Range },
) {
  sheet.insertChart(
    sheet
      .newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(tables.totals)
      .setNumHeaders(1)
      .setPosition(row, 1, 0, 0)
      .setOption("title", `${label}: total scores received, by tournament`)
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
      .addRange(tables.categories)
      .setNumHeaders(1)
      .setPosition(row, 8, 0, 0)
      .setOption("title", `${label}: category scores received`)
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
 * rebuild the Division Report tab from scratch: a heading and two charts per group, then the tables they are drawn from
 * written as values, so it only changes when Refresh Club Statistics runs
 *
 * @param events  all events, in date order
 * @param scores  every score at an included tournament
 * @returns a one-line summary
 */
function _refreshDivisionReport_(events: EventRecord[], scores: CountedScore[]): string {
  const groups = _divisionGroups_(events, scores);
  const sheet = _getDivisionReportSheet_();
  for (const chart of sheet.getCharts()) sheet.removeChart(chart);
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).breakApart();
  sheet.clear();

  const tablesTop = DIVISION_REPORT_FIRST_ROW + groups.length * DIVISION_REPORT_GROUP_ROWS;
  const widest = Math.max(0, ...groups.map((g) => g.tournaments.length));
  _fitSheet_(
    sheet,
    tablesTop + groups.length * _divisionReportTableRows_(),
    Math.max(CLUB_REPORT_INFO_WIDTH, 1 + widest),
  );
  sheet
    .getRange(INFO_ROW, 1, 1, CLUB_REPORT_INFO_WIDTH)
    .merge()
    .setValue(_divisionReportInfo_(new Date()))
    .setWrap(true)
    .setVerticalAlignment("top");

  groups.forEach((group, i) => {
    const heading = DIVISION_REPORT_FIRST_ROW + i * DIVISION_REPORT_GROUP_ROWS;
    sheet.getRange(heading, 1, 1, 2).setValues([[group.label, _divisionSummary_(group)]]);
    sheet.getRange(heading, 1).setFontWeight("bold").setFontSize(12);
    sheet.getRange(heading, 2, 1, 6).merge();
    const tables = _writeDivisionTables_(sheet, tablesTop + i * _divisionReportTableRows_(), group);
    if (group.tournaments.length > 0) _insertDivisionCharts_(sheet, heading + 1, group.label, tables);
  });

  const noDivision =
    events.filter((e) => e.division === "" && scores.some((s) => s.counts && s.tournament === e.spiritFormUrl)).length;
  return noDivision === 0 ? "" : `${noDivision} tournament(s) with counted scores have no Div`;
}
