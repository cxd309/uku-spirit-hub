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
  + "To Refresh run \"Scan for files\"";

/**
 * responses tab columns, in order
 * keys are the names used in code
 * values are the header text
 * columns listed in RESPONSE_FORMULAS are formulas, the rest hold ResponseRecord values
 */
const RESPONSE_HEADERS = Object.freeze(
  /** @type {const} */ ({
    fileId: "File ID",
    sourceRow: "Source Row",
    scorer: "Scoring Team",
    scorerClub: "Scoring Club",
    receiver: "Receiving Team",
    receiverClub: "Receiving Club",
    tournament: "Tournament",
    rules: "Rules",
    fouls: "Fouls",
    fairMindedness: "Fair-Mindedness",
    attitude: "Attitude",
    communication: "Communication",
    comment: "Comment",
  }),
);

/**
 * formula columns of the Responses tab:
 * for each, a function building that column's formula for a given sheet row
 * every row gets its own formula, so sorting or filtering the tab never breaks them
 *
 * @type {Readonly<Partial<Record<keyof typeof RESPONSE_HEADERS, function(number): string>>>}
 */
const RESPONSE_FORMULAS = Object.freeze({
  scorerClub: (/** @type {number} */ row) => _responseClubFormula_(row, "scorer"),
  receiverClub: (/** @type {number} */ row) => _responseClubFormula_(row, "receiver"),
  tournament: _tournamentFormula_,
});

/**
 * keys of ResponseRecord in column order
 */
const RESPONSE_KEYS = /** @type {(keyof typeof RESPONSE_HEADERS)[]} */ (Object.keys(RESPONSE_HEADERS));

/**
 * one row of the responses tab aone team's scores for one opponent at one event
 *
 * @typedef {Object} ResponseRecord
 * @property {string} fileId          file ID of the event (links to the Events tab)
 * @property {number} sourceRow       row in the source breakdown tab
 * @property {string} scorer          team giving the score
 * @property {string} receiver        team receiving the score
 * @property {number} rules           rules Knowledge and Use, 0–4
 * @property {number} fouls           fouls and Body Contact, 0–4
 * @property {number} fairMindedness  fair-Mindedness, 0–4
 * @property {number} attitude        positive Attitude and Self-Control, 0–4
 * @property {number} communication   communication, 0–4
 * @property {string} comment         comment; "" if none
 */

/**
 * @param {unknown[]} row  values of one data row, in RESPONSE_KEYS order
 * @returns {ResponseRecord} the record
 */
function _responseFromRow_(row) {
  /** @param {keyof typeof RESPONSE_HEADERS} key */
  const cell = (key) => row[RESPONSE_KEYS.indexOf(key)];
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
 * @param {ResponseRecord}  response  the record
 * @param {number}          row       1-based sheet row it will be written to
 * @returns {unknown[]} values in RESPONSE_KEYS order
 */
function _responseToRow_(response, row) {
  return RESPONSE_KEYS.map((key) => {
    const formula = RESPONSE_FORMULAS[key];
    return formula ? formula(row) : response[/** @type {keyof ResponseRecord} */ (key)];
  });
}
/**
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} the responses tab, created on first use
 */
function _getResponsesSheet_() {
  /** @type {readonly string[]} */
  const headers = Object.values(RESPONSE_HEADERS);
  return _getOrCreateSheet_(SpreadsheetApp.getActiveSpreadsheet(), RESPONSES_SHEET, headers, RESPONSES_INFO).sheet;
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet  the Responses tab
 * @returns {ResponseRecord[]} every response on the tab
 */
function _readResponses_(sheet) {
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
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet      the Responses tab
 * @param {ResponseRecord[]}                   responses  all responses, in display order
 */
function _writeResponses_(sheet, responses) {
  _fitSheet_(sheet, HEADER_ROW + responses.length, RESPONSE_KEYS.length);
  if (responses.length === 0) return;
  sheet
    .getRange(DATA_ROW, 1, responses.length, RESPONSE_KEYS.length)
    .setValues(responses.map((r, i) => _responseToRow_(r, i + DATA_ROW)));
}

/**
 * tournament formula for one Responses row
 * looks the event up by file id on the Events tab
 *
 * @param {number} row  1-based sheet row the formula is for
 * @returns {string} the formula
 */
function _tournamentFormula_(row) {
  const id = `$${_columnLetter_(RESPONSE_KEYS.indexOf("fileId") + 1)}${row}`;
  /** @param {keyof typeof EVENT_HEADERS} key */
  const events = (key) => _columnBelowHeader_(EVENTS_SHEET, EVENT_KEYS.indexOf(key) + 1);
  return `=IFERROR(XLOOKUP(${id}, ${events("fileId")}, ${events("tournament")}), "(unknown event)")`;
}

/**
 * club formula for one side of one Responses row
 * looks the team up on the Teams tab and returns its Club
 * blank if the team is not on the Teams tab
 *   e.g. foreign teams at international events
 *
 * @param {number}                 row   1-based sheet row the formula is for
 * @param {"scorer" | "receiver"}  side  which team on the row
 * @returns {string} the formula
 */
function _responseClubFormula_(row, side) {
  const team = `$${_columnLetter_(RESPONSE_KEYS.indexOf(side) + 1)}${row}`;
  /** @param {keyof typeof TEAM_HEADERS} key */
  const teams = (key) => _columnBelowHeader_(TEAMS_SHEET, TEAM_KEYS.indexOf(key) + 1);
  return `=IFERROR(XLOOKUP(${team}, ${teams("team")}, ${teams("club")}), "")`;
}

/**
 * responses in the date order of their events, oldest first
 * then by tournament, event, and source row, so each event's responses stay in their original order
 * pure, returns a new array
 *
 * @param {ResponseRecord[]} responses  responses in any order
 * @param {EventRecord[]}    events     events, used to find each response's date and tournament
 * @returns {ResponseRecord[]} responses in date order
 */
function _sortResponses_(responses, events) {
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
