/**
 * header of the Results tab of a published file, as the Schedule Manager's pivot table wrote it
 * the team column has no header, and Count has a trailing space
 */
const PUBLISHED_RESULTS_HEADERS = Object.freeze([
  "",
  "Count ",
  "Rules",
  "Fouls",
  "Fairness",
  "Attitude",
  "Comms",
  "Total",
]);

/**
 * header of the Breakdown tab of a published file: the form's own questions, without Timestamp and Repeat
 */
const PUBLISHED_BREAKDOWN_HEADERS = Object.freeze([
  FORM_HEADERS.yourTeam,
  FORM_HEADERS.opponentTeam,
  FORM_HEADERS.rules,
  FORM_HEADERS.fouls,
  FORM_HEADERS.fairMindedness,
  FORM_HEADERS.attitude,
  FORM_HEADERS.communication,
  FORM_HEADERS.comments,
]);

/**
 * colours of a published file, as the Schedule Manager's formatting used them
 */
const PUBLISHED_COLOURS = Object.freeze({
  black: "#000000",
  white: "#ffffff",
  red: "#ff0000",
  headerGrey: "#bdbdbd",
  rowGrey: "#f0f0f0",
  lowCategory: "#ea9999",
  lowTotal: "#f45b5b",
  monitoringTotal: "#fcb375",
  zero: "#f4c7c3",
  four: "#b7e1cd",
});

/**
 * the two tables of a published file, as values
 */
interface PublishedTables {
  /** header, then one row per receiving team: team, count, five category averages, average total */
  results: (string | number)[][];
  /** header, then one row per counted form: scoring team, receiving team, five scores, comment */
  breakdown: (string | number)[][];
}

/**
 * the Results and Breakdown tables of one tournament, from its counted forms
 * pure, never modifies its arguments
 *
 * Results: each receiving team's number of scores and averages, highest Total first,
 * ties in reverse order of team name, as the Schedule Manager's pivot table ordered them
 * Breakdown: every counted form, by receiving team, then scoring team
 *
 * @param responses  the tournament's responses, only counted ones are used
 * @returns the two tables, each with its header row
 */
function _publishedTables_(responses: ResponseRecord[]): PublishedTables {
  const counted = responses.filter((r) => r.counted);
  const byTeam: Map<string, ResponseRecord[]> = new Map();
  for (const r of counted) byTeam.set(r.receiver, [...(byTeam.get(r.receiver) ?? []), r]);

  const average = (forms: ResponseRecord[], value: (r: ResponseRecord) => number) =>
    forms.reduce((sum, r) => sum + value(r), 0) / forms.length;
  const results = [...byTeam].map(([team, forms]) => [
    team,
    forms.length,
    ...SCORE_KEYS.map((key) => average(forms, (r) => r[key])),
    average(forms, _responseTotal_),
  ]).sort((a, b) => (b[7] as number) - (a[7] as number) || String(b[0]).localeCompare(String(a[0])));

  const breakdown = [...counted]
    .sort((a, b) => a.receiver.localeCompare(b.receiver) || a.scorer.localeCompare(b.scorer))
    .map((r) => [r.scorer, r.receiver, ...SCORE_KEYS.map((key) => r[key]), r.comment]);

  return {
    results: [[...PUBLISHED_RESULTS_HEADERS], ...results],
    breakdown: [[...PUBLISHED_BREAKDOWN_HEADERS], ...breakdown],
  };
}

/**
 * a colour for the Sheets API
 *
 * @param hex  e.g. "#ea9999"
 * @returns red, green and blue from 0 to 1
 */
function _apiColour_(hex: string): GoogleAppsScript.Sheets.Schema.Color {
  const channel = (i: number) => parseInt(hex.slice(1 + 2 * i, 3 + 2 * i), 16) / 255;
  return { red: channel(0), green: channel(1), blue: channel(2) };
}

/**
 * a cell for the Sheets API's updateCells: its value and its whole format
 *
 * @param value   text or number
 * @param format  the cell's format
 * @returns the cell
 */
function _apiCell_(
  value: string | number,
  format: GoogleAppsScript.Sheets.Schema.CellFormat,
): GoogleAppsScript.Sheets.Schema.CellData {
  const userEnteredValue = typeof value === "number" ? { numberValue: value } : { stringValue: value };
  return { userEnteredValue, userEnteredFormat: format };
}

/**
 * a conditional format rule for the Sheets API, colouring cells in some columns of a tab
 *
 * @param sheetId    the tab
 * @param columns    first and last 0-based column, inclusive
 * @param condition  e.g. NUMBER_LESS_THAN_EQ
 * @param values     the condition's values
 * @param colour     the background
 * @returns the request
 */
function _apiColourRule_(
  sheetId: number,
  columns: [number, number],
  condition: string,
  values: number[],
  colour: string,
): GoogleAppsScript.Sheets.Schema.Request {
  return {
    addConditionalFormatRule: {
      rule: {
        ranges: [{ sheetId, startRowIndex: 1, startColumnIndex: columns[0], endColumnIndex: columns[1] + 1 }],
        booleanRule: {
          condition: { type: condition, values: values.map((v) => ({ userEnteredValue: String(v) })) },
          format: { backgroundColor: _apiColour_(colour) },
        },
      },
      index: 0,
    },
  };
}

/**
 * the tab of a published file the requests are written for
 */
interface PublishedTab {
  /** the tab's ID, as the Sheets API gives it */
  sheetId: number;
  /** whether the tab has to be added first */
  isNew: boolean;
  /** how many conditional format rules it has now, all removed before the new ones are added */
  ruleCount: number;
}

/**
 * every Sheets API request that turns two tabs into a published Results and Breakdown, as today's files look
 * pure, no google calls
 * each tab is resized to the table, every cell's value and format is written, so nothing from before is left,
 * then the column widths and the colour rules are set
 *
 * @param tables        from _publishedTables_
 * @param titles        tab names: "<Event> SOTG Results" and "<Event> SOTG Breakdown"
 * @param resultsTab    the tab for Results
 * @param breakdownTab  the tab for Breakdown
 * @returns the requests, in order
 */
function _publishRequests_(
  tables: PublishedTables,
  titles: { results: string; breakdown: string },
  resultsTab: PublishedTab,
  breakdownTab: PublishedTab,
): GoogleAppsScript.Sheets.Schema.Request[] {
  const colour = (name: keyof typeof PUBLISHED_COLOURS) => _apiColour_(PUBLISHED_COLOURS[name]);
  const thin = { style: "SOLID", color: colour("black") };
  const requests: GoogleAppsScript.Sheets.Schema.Request[] = [];

  const prepare = (tab: PublishedTab, title: string, index: number, rows: number) => {
    const gridProperties = { rowCount: rows, columnCount: PUBLISHED_RESULTS_HEADERS.length, frozenRowCount: 0 };
    if (tab.isNew) {
      requests.push({ addSheet: { properties: { sheetId: tab.sheetId, title, index, gridProperties } } });
    } else {
      requests.push({
        updateSheetProperties: {
          properties: { sheetId: tab.sheetId, title, index, gridProperties },
          fields: "title,index,gridProperties(rowCount,columnCount,frozenRowCount)",
        },
      });
    }
    for (let i = tab.ruleCount - 1; i >= 0; i--) {
      requests.push({ deleteConditionalFormatRule: { sheetId: tab.sheetId, index: i } });
    }
  };
  const widths = (sheetId: number, sizes: number[]) =>
    sizes.forEach((pixelSize, i) =>
      requests.push({
        updateDimensionProperties: {
          range: { sheetId, dimension: "COLUMNS", startIndex: i, endIndex: i + 1 },
          properties: { pixelSize },
          fields: "pixelSize",
        },
      })
    );

  // Results: every cell bold with thin borders, numbers to 2 places, the count as a whole number
  prepare(resultsTab, titles.results, 0, tables.results.length);
  const resultRows = tables.results.map((row, r) => ({
    values: row.map((value, c) =>
      _apiCell_(value, {
        textFormat: { bold: true },
        backgroundColor: colour("white"),
        borders: { top: thin, bottom: thin, left: thin, right: thin },
        horizontalAlignment: c === 0 ? "LEFT" : "RIGHT",
        ...(r > 0 && c > 0 ? { numberFormat: { type: "NUMBER", pattern: c === 1 ? "0" : "0.00" } } : {}),
      })
    ),
  }));
  requests.push({
    updateCells: {
      start: { sheetId: resultsTab.sheetId, rowIndex: 0, columnIndex: 0 },
      rows: resultRows,
      fields: "userEnteredValue,userEnteredFormat",
    },
  });
  widths(resultsTab.sheetId, [150, 100, 100, 100, 100, 100, 100, 100]);
  requests.push(
    _apiColourRule_(resultsTab.sheetId, [2, 6], "NUMBER_LESS_THAN_EQ", [1.8], PUBLISHED_COLOURS.lowCategory),
    _apiColourRule_(resultsTab.sheetId, [7, 7], "NUMBER_LESS_THAN_EQ", [8], PUBLISHED_COLOURS.lowTotal),
    _apiColourRule_(resultsTab.sheetId, [7, 7], "NUMBER_BETWEEN", [8.01, 9], PUBLISHED_COLOURS.monitoringTotal),
  );

  // Breakdown: grey header, alternating rows, a red line where the receiving team changes, comments wrapped
  prepare(breakdownTab, titles.breakdown, 1, tables.breakdown.length);
  const commentColumn = PUBLISHED_BREAKDOWN_HEADERS.length - 1;
  const breakdownRows = tables.breakdown.map((row, r) => {
    const newTeam = r > 1 && row[1] !== tables.breakdown[r - 1][1];
    const background = r === 0 ? colour("headerGrey") : r % 2 === 1 ? colour("white") : colour("rowGrey");
    return {
      values: row.map((value, c) =>
        _apiCell_(value, {
          backgroundColor: background,
          horizontalAlignment: typeof value === "number" ? "RIGHT" : "LEFT",
          wrapStrategy: c === commentColumn ? "WRAP" : "OVERFLOW_CELL",
          ...(newTeam ? { borders: { top: { style: "SOLID_MEDIUM", color: colour("red") } } } : {}),
        })
      ),
    };
  });
  requests.push({
    updateCells: {
      start: { sheetId: breakdownTab.sheetId, rowIndex: 0, columnIndex: 0 },
      rows: breakdownRows,
      fields: "userEnteredValue,userEnteredFormat",
    },
  });
  widths(breakdownTab.sheetId, [100, 100, 100, 100, 100, 100, 100, 880]);
  requests.push(
    _apiColourRule_(breakdownTab.sheetId, [2, 6], "NUMBER_EQ", [0], PUBLISHED_COLOURS.zero),
    _apiColourRule_(breakdownTab.sheetId, [2, 6], "NUMBER_EQ", [4], PUBLISHED_COLOURS.four),
  );
  return requests;
}
