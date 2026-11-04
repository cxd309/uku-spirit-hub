/**
 * name of the responses tab
 */
const RESPONSES_SHEET = "Results";

/**
 * text for the info row above the Responses table
 * what the tab is, what to edit, how it refreshes
 */
const RESPONSES_INFO = "All spirit scores across all tournaments\n\n"
  + "Each form is matched to a game on the Games tab. Counted is unticked for forms that do not count: "
  + "Match says Duplicate (more forms than games, the one with a comment then the latest counts), "
  + "Self score or No such game. No schedule means the tournament's schedule was not imported, so nothing is matched\n\n"
  + "DO NOT EDIT, this is all regenerated on refresh. Filters are kept through a refresh, sorting is put back to the hub's order\n\n"
  + "Report Team is the Receiving Team without trailing brackets, e.g. Birmingham 1 (WW4) is Birmingham 1, for the Club Report\n\n"
  + "To Refresh run \"Refresh Results and Issues\"";

/**
 * responses tab columns, in order
 * keys are the names used in code
 * values are the header text
 * columns listed in RESPONSE_FORMULAS are formulas, the rest hold ResponseRecord values
 */
const RESPONSE_HEADERS = Object.freeze(
  {
    spiritFormUrl: "Spirit Form URL",
    date: "Date",
    tournament: "Tournament",
    sourceRow: "Source Row",
    timestamp: "Timestamp",
    scorer: "Scoring Team",
    scorerClub: "Scoring Club",
    receiver: "Receiving Team",
    receiverClub: "Receiving Club",
    rules: "Rules",
    fouls: "Fouls",
    fairMindedness: "Fair-Mindedness",
    attitude: "Attitude",
    communication: "Communication",
    total: "Total",
    comment: "Comment",
    game: "Game",
    counted: "Counted",
    match: "Match",
    spiritAward: "Spirit Award",
    countsForAward: "Counts for Award",
    reportTeam: "Report Team",
  } as const,
);

/**
 * trailing bracket groups removed from a receiving team for its Report Team
 * e.g. Birmingham 1 (WW4) and Birmingham 1 (WW2) (3) are both Birmingham 1
 * the same syntax as a Sheets REGEXREPLACE
 */
const REPORT_TEAM_PATTERN = "(\\s*\\([^)]*\\))+$";

/**
 * formula columns of the Responses tab:
 * for each, a function building that column's formula for a given sheet row
 * every row gets its own formula, so sorting or filtering the tab never breaks them
 */
const RESPONSE_FORMULAS: Readonly<
  Partial<Record<keyof typeof RESPONSE_HEADERS, (row: number) => string>>
> = Object.freeze({
  date: (row: number) => {
    const id = `$${_columnLetter_(RESPONSE_KEYS.indexOf("spiritFormUrl") + 1)}${row}`;
    const events = (key: keyof typeof EVENT_HEADERS) => _columnBelowHeader_(EVENTS_SHEET, EVENT_KEYS.indexOf(key) + 1);
    return `=IFERROR(XLOOKUP(${id}, ${events("spiritFormUrl")}, ${events("startDate")}), "")`;
  },
  scorerClub: (row: number) => _responseClubFormula_(row, "scorer"),
  receiverClub: (row: number) => _responseClubFormula_(row, "receiver"),
  tournament: _tournamentFormula_,
  spiritAward: (row: number) => {
    const id = `$${_columnLetter_(RESPONSE_KEYS.indexOf("spiritFormUrl") + 1)}${row}`;
    const events = (key: keyof typeof EVENT_HEADERS) => _columnBelowHeader_(EVENTS_SHEET, EVENT_KEYS.indexOf(key) + 1);
    const lookup = (key: keyof typeof EVENT_HEADERS) => `XLOOKUP(${id}, ${events("spiritFormUrl")}, ${events(key)})`;
    const international = _readAwardSettings_().includeInternational ? "TRUE" : `NOT(${lookup("international")})`;
    return `=IFERROR(AND(${lookup("includeResults")}, ${international}), FALSE)`;
  },
  countsForAward: (row: number) => {
    const cell = (key: keyof typeof RESPONSE_HEADERS) => `$${_columnLetter_(RESPONSE_KEYS.indexOf(key) + 1)}${row}`;
    return `=AND(${cell("spiritAward")}, ${cell("counted")}, ${cell("scorerClub")}<>${cell("receiverClub")})`;
  },
  reportTeam: (row: number) => {
    const receiver = `$${_columnLetter_(RESPONSE_KEYS.indexOf("receiver") + 1)}${row}`;
    return `=TRIM(REGEXREPLACE(${receiver}&"", "${REPORT_TEAM_PATTERN}", ""))`;
  },
  total: (row: number) => {
    const cells = SCORE_KEYS.map((key) => `$${_columnLetter_(RESPONSE_KEYS.indexOf(key) + 1)}${row}`);
    return `=${cells.join("+")}`;
  },
});

/**
 * keys of ResponseRecord in column order
 */
const RESPONSE_KEYS = Object.keys(RESPONSE_HEADERS) as (keyof typeof RESPONSE_HEADERS)[];

/**
 * one row of the responses tab: one team's scores for one opponent at one event
 */
interface ResponseRecord {
  /** link to the tournament's Form Responses (links to the Events tab) */
  spiritFormUrl: string;
  /** row in the Form Responses' responses tab */
  sourceRow: number;
  /** when the form was submitted, null if blank */
  timestamp: Date | null;
  /** team giving the score */
  scorer: string;
  /** team receiving the score */
  receiver: string;
  /** rules Knowledge and Use, 0–4 */
  rules: number;
  /** fouls and Body Contact, 0–4 */
  fouls: number;
  /** fair-Mindedness, 0–4 */
  fairMindedness: number;
  /** positive Attitude and Self-Control, 0–4 */
  attitude: number;
  /** communication, 0–4 */
  communication: number;
  /** comment; "" if none */
  comment: string;
  /** game code it was matched to, "" if none */
  game: string;
  /** whether this form counts, false for duplicates, self scores and forms with no game */
  counted: boolean;
  /** one of MATCH_RESULTS */
  match: string;
}

/**
 * @param row  values of one data row, in RESPONSE_KEYS order
 * @returns the record
 */
function _responseFromRow_(row: unknown[]): ResponseRecord {
  const cell = (key: keyof typeof RESPONSE_HEADERS) => row[RESPONSE_KEYS.indexOf(key)];
  return {
    spiritFormUrl: String(cell("spiritFormUrl")),
    sourceRow: Number(cell("sourceRow")),
    timestamp: cell("timestamp") instanceof Date ? cell("timestamp") as Date : null,
    scorer: String(cell("scorer")),
    receiver: String(cell("receiver")),
    rules: Number(cell("rules")),
    fouls: Number(cell("fouls")),
    fairMindedness: Number(cell("fairMindedness")),
    attitude: Number(cell("attitude")),
    communication: Number(cell("communication")),
    comment: String(cell("comment")),
    game: String(cell("game")),
    counted: cell("counted") === true,
    match: String(cell("match")),
  };
}

/**
 * @param response  the record
 * @returns values in RESPONSE_KEYS order, "" in formula columns (filled by _fillFormulaColumns_)
 */
function _responseToRow_(response: ResponseRecord): unknown[] {
  return RESPONSE_KEYS.map((key) => {
    if (RESPONSE_FORMULAS[key]) return "";
    const value = response[key as keyof ResponseRecord];
    return value === null ? "" : value;
  });
}
/**
 * the parts of a response that say which teams played at which tournament
 */
type ResponseTeams = Pick<ResponseRecord, "spiritFormUrl" | "scorer" | "receiver">;

/**
 * read only the Spirit Form URL, Scoring Team and Receiving Team columns of the Responses tab
 * plain values, so this never waits for the tab's formulas to calculate
 *
 * @param sheet  the Responses tab
 * @returns who played at which tournament, one per response
 */
function _readResponseTeams_(sheet: GoogleAppsScript.Spreadsheet.Sheet): ResponseTeams[] {
  const rowCount = sheet.getLastRow() - HEADER_ROW;
  if (rowCount < 1) return [];
  const column = (key: keyof typeof RESPONSE_HEADERS) =>
    sheet.getRange(DATA_ROW, RESPONSE_KEYS.indexOf(key) + 1, rowCount, 1).getValues().map((row) => String(row[0]));
  const urls = column("spiritFormUrl");
  const scorers = column("scorer");
  const receivers = column("receiver");
  return urls
    .map((spiritFormUrl, i) => ({ spiritFormUrl, scorer: scorers[i], receiver: receivers[i] }))
    .filter((r) => r.spiritFormUrl !== "");
}

/**
 * @returns the responses tab, created on first use
 */
function _getResponsesSheet_(): GoogleAppsScript.Spreadsheet.Sheet {
  const headers: readonly string[] = Object.values(RESPONSE_HEADERS);
  return _getOrCreateSheet_(SpreadsheetApp.getActiveSpreadsheet(), RESPONSES_SHEET, headers, RESPONSES_INFO);
}

/**
 * @param sheet  the Responses tab
 * @returns every response on the tab
 */
function _readResponses_(sheet: GoogleAppsScript.Spreadsheet.Sheet): ResponseRecord[] {
  const rowCount = sheet.getLastRow() - HEADER_ROW;
  if (rowCount < 1) return [];
  return sheet
    .getRange(DATA_ROW, 1, rowCount, RESPONSE_KEYS.length)
    .getValues()
    .map(_responseFromRow_)
    .filter((r) => r.spiritFormUrl !== "");
}

/**
 * replace the contents of the Responses tab, resizing it to fit
 * values are written first, then each formula column is filled in one call
 *
 * @param sheet      the Responses tab
 * @param responses  all responses, in display order
 */
function _writeResponses_(sheet: GoogleAppsScript.Spreadsheet.Sheet, responses: ResponseRecord[]) {
  _timed_("results values", () => _writeTable_(sheet, responses.map(_responseToRow_), RESPONSE_KEYS.length));
  _timed_("results formulas", () => _fillFormulaColumns_(sheet, RESPONSE_KEYS, RESPONSE_FORMULAS, responses.length));
  if (responses.length > 0) {
    sheet.getRange(DATA_ROW, RESPONSE_KEYS.indexOf("date") + 1, responses.length, 1).setNumberFormat("yyyy-mm-dd");
    sheet
      .getRange(DATA_ROW, RESPONSE_KEYS.indexOf("timestamp") + 1, responses.length, 1)
      .setNumberFormat("yyyy-mm-dd hh:mm:ss");
    sheet
      .getRange(DATA_ROW, RESPONSE_KEYS.indexOf("counted") + 1, responses.length, 1)
      .setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
  }
}

/**
 * tournament formula for one Responses row
 * looks the event up by its Spirit Form URL on the Events tab
 *
 * @param row  1-based sheet row the formula is for
 * @returns the formula
 */
function _tournamentFormula_(row: number): string {
  const id = `$${_columnLetter_(RESPONSE_KEYS.indexOf("spiritFormUrl") + 1)}${row}`;
  const events = (key: keyof typeof EVENT_HEADERS) => _columnBelowHeader_(EVENTS_SHEET, EVENT_KEYS.indexOf(key) + 1);
  return `=IFERROR(XLOOKUP(${id}, ${events("spiritFormUrl")}, ${events("eventName")}), "(unknown event)")`;
}

/**
 * club formula for one side of one Responses row
 * looks the team up on the Teams tab and returns its Club
 * blank if the team is not on the Teams tab
 *   e.g. foreign teams at international events
 *
 * @param row   1-based sheet row the formula is for
 * @param side  which team on the row
 * @returns the formula
 */
function _responseClubFormula_(row: number, side: "scorer" | "receiver"): string {
  const team = `$${_columnLetter_(RESPONSE_KEYS.indexOf(side) + 1)}${row}`;
  const teams = (key: keyof typeof TEAM_HEADERS) => _columnBelowHeader_(TEAMS_SHEET, TEAM_KEYS.indexOf(key) + 1);
  return `=IFERROR(XLOOKUP(${team}, ${teams("team")}, ${teams("club")}), "")`;
}

/**
 * responses in the date order of their events, oldest first
 * then by tournament, event, and source row, so each event's responses stay in their original order
 * pure, returns a new array
 *
 * @param responses  responses in any order
 * @param events     events, used to find each response's date and tournament
 * @returns responses in date order
 */
function _sortResponses_(responses: ResponseRecord[], events: EventRecord[]): ResponseRecord[] {
  const eventById = new Map(events.map((e) => [e.spiritFormUrl, e]));
  return [...responses].sort((a, b) => {
    const ea = eventById.get(a.spiritFormUrl);
    const eb = eventById.get(b.spiritFormUrl);
    return _compareDates_(ea?.startDate ?? null, eb?.startDate ?? null)
      || (ea?.eventName ?? "").localeCompare(eb?.eventName ?? "")
      || a.spiritFormUrl.localeCompare(b.spiritFormUrl)
      || a.sourceRow - b.sourceRow;
  });
}
