/**
 * name of the responses tab
 */
const RESPONSES_SHEET = "Results";

/**
 * text for the info row above the Responses table
 * what the tab is, what to edit, how it refreshes
 */
const RESPONSES_INFO = "All spirit scores across all tournaments\n\n"
  + "DO NOT EDIT, this is all regenerated on refresh. Sort and filter is safe but will be overwritten on refresh\n\n"
  + "To Refresh run \"Refresh Results\"";

/**
 * responses tab columns, in order
 * keys are the names used in code
 * values are the header text
 * columns listed in RESPONSE_FORMULAS are formulas, the rest hold ResponseRecord values
 */
const RESPONSE_HEADERS = Object.freeze(
  {
    fileId: "File ID",
    sourceRow: "Source Row",
    scorer: "Scoring Team",
    scorerClub: "Scoring Club",
    receiver: "Receiving Team",
    receiverClub: "Receiving Club",
    tournament: "Tournament",
    included: "Included",
    countsForAward: "Counts for Award",
    rules: "Rules",
    fouls: "Fouls",
    fairMindedness: "Fair-Mindedness",
    attitude: "Attitude",
    communication: "Communication",
    total: "Total",
    comment: "Comment",
  } as const,
);

/**
 * formula columns of the Responses tab:
 * for each, a function building that column's formula for a given sheet row
 * every row gets its own formula, so sorting or filtering the tab never breaks them
 */
const RESPONSE_FORMULAS: Readonly<
  Partial<Record<keyof typeof RESPONSE_HEADERS, (row: number) => string>>
> = Object.freeze({
  scorerClub: (row: number) => _responseClubFormula_(row, "scorer"),
  receiverClub: (row: number) => _responseClubFormula_(row, "receiver"),
  tournament: _tournamentFormula_,
  included: (row: number) => {
    const id = `$${_columnLetter_(RESPONSE_KEYS.indexOf("fileId") + 1)}${row}`;
    const events = (key: keyof typeof EVENT_HEADERS) => _columnBelowHeader_(EVENTS_SHEET, EVENT_KEYS.indexOf(key) + 1);
    return `=IFERROR(XLOOKUP(${id}, ${events("fileId")}, ${events("include")}), FALSE)`;
  },
  countsForAward: (row: number) => {
    const cell = (key: keyof typeof RESPONSE_HEADERS) => `$${_columnLetter_(RESPONSE_KEYS.indexOf(key) + 1)}${row}`;
    return `=AND(${cell("included")}, ${cell("scorerClub")}<>${cell("receiverClub")})`;
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
 * one row of the responses tab aone team's scores for one opponent at one event
 */
interface ResponseRecord {
  /** file ID of the event (links to the Events tab) */
  fileId: string;
  /** row in the source breakdown tab */
  sourceRow: number;
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
}

/**
 * @param row  values of one data row, in RESPONSE_KEYS order
 * @returns the record
 */
function _responseFromRow_(row: unknown[]): ResponseRecord {
  const cell = (key: keyof typeof RESPONSE_HEADERS) => row[RESPONSE_KEYS.indexOf(key)];
  return {
    fileId: String(cell("fileId")),
    sourceRow: Number(cell("sourceRow")),
    scorer: String(cell("scorer")),
    receiver: String(cell("receiver")),
    rules: Number(cell("rules")),
    fouls: Number(cell("fouls")),
    fairMindedness: Number(cell("fairMindedness")),
    attitude: Number(cell("attitude")),
    communication: Number(cell("communication")),
    comment: String(cell("comment")),
  };
}

/**
 * @param response  the record
 * @param row       1-based sheet row it will be written to
 * @returns values in RESPONSE_KEYS order
 */
function _responseToRow_(response: ResponseRecord, row: number): unknown[] {
  return RESPONSE_KEYS.map((key) => {
    const formula = RESPONSE_FORMULAS[key];
    return formula ? formula(row) : response[key as keyof ResponseRecord];
  });
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
    .filter((r) => r.fileId !== "");
}

/**
 * replace the contents of the Responses tab, resizing it to fit
 *
 * @param sheet      the Responses tab
 * @param responses  all responses, in display order
 */
function _writeResponses_(sheet: GoogleAppsScript.Spreadsheet.Sheet, responses: ResponseRecord[]) {
  _writeTable_(sheet, responses.map((r, i) => _responseToRow_(r, i + DATA_ROW)), RESPONSE_KEYS.length);
}

/**
 * tournament formula for one Responses row
 * looks the event up by file id on the Events tab
 *
 * @param row  1-based sheet row the formula is for
 * @returns the formula
 */
function _tournamentFormula_(row: number): string {
  const id = `$${_columnLetter_(RESPONSE_KEYS.indexOf("fileId") + 1)}${row}`;
  const events = (key: keyof typeof EVENT_HEADERS) => _columnBelowHeader_(EVENTS_SHEET, EVENT_KEYS.indexOf(key) + 1);
  return `=IFERROR(XLOOKUP(${id}, ${events("fileId")}, ${events("tournament")}), "(unknown event)")`;
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
  const eventById = new Map(events.map((e) => [e.fileId, e]));
  return [...responses].sort((a, b) => {
    const ea = eventById.get(a.fileId);
    const eb = eventById.get(b.fileId);
    return _compareDates_(ea?.date ?? null, eb?.date ?? null)
      || (ea?.tournament ?? "").localeCompare(eb?.tournament ?? "")
      || a.fileId.localeCompare(b.fileId)
      || a.sourceRow - b.sourceRow;
  });
}
