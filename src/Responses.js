/**
 * name of the responses tab
 */
const RESPONSES_SHEET = "Responses";

/**
 * header of the formula column after the data columns
 * its values are looked up from the Events tab by a single array formula
 */
const RESPONSE_TOURNAMENT_HEADER = "Tournament";

/**
 * responses tab columns, in order
 * keys are the names used in code
 * values are the header text
 */
const RESPONSE_HEADERS = Object.freeze(
  /** @type {const} */ ({
    fileId: "File ID",
    sourceRow: "Source Row",
    scorer: "Scoring Team",
    receiver: "Receiving Team",
    rules: "Rules",
    fouls: "Fouls",
    fairMindedness: "Fair-Mindedness",
    attitude: "Attitude",
    communication: "Communication",
    comment: "Comment",
  }),
);

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
 * @param {ResponseRecord} response  the record
 * @returns {unknown[]} values in RESPONSE_KEYS order
 */
function _responseToRow_(response) {
  return RESPONSE_KEYS.map((key) => response[key]);
}

/**
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} the responses tab, created on first use
 */
function _getResponsesSheet_() {
  /** @type {readonly string[]} */
  const headers = [...Object.values(RESPONSE_HEADERS), RESPONSE_TOURNAMENT_HEADER];
  return _getOrCreateSheet_(SpreadsheetApp.getActiveSpreadsheet(), RESPONSES_SHEET, headers).sheet;
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet  the Responses tab
 * @returns {ResponseRecord[]} every response on the tab
 */
function _readResponses_(sheet) {
  const rowCount = sheet.getLastRow() - 1;
  if (rowCount < 1) return [];
  return sheet
    .getRange(2, 1, rowCount, RESPONSE_KEYS.length)
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
  const tournamentColumn = RESPONSE_KEYS.length + 1;
  _fitSheet_(sheet, 1 + responses.length, tournamentColumn);
  if (responses.length > 0) {
    sheet.getRange(2, 1, responses.length, RESPONSE_KEYS.length).setValues(responses.map(_responseToRow_));
  }
  sheet.getRange(2, tournamentColumn).setFormula(_tournamentFormula_());
}

/**
 * build the array formula that fills the Tournament column:
 * for each response, the event's Name Override if set, otherwise its Default Name
 *
 * column letters are derived from RESPONSE_KEYS and EVENT_KEYS,
 * so reordering either header list keeps the formula correct.
 *
 * @returns {string} the formula for row 2 of the Tournament column.
 */
function _tournamentFormula_() {
  const id = _columnLetter_(RESPONSE_KEYS.indexOf("fileId") + 1);
  /** @param {keyof typeof EVENT_HEADERS} key */
  const events = (key) => {
    const letter = _columnLetter_(EVENT_KEYS.indexOf(key) + 1);
    return `'${EVENTS_SHEET}'!${letter}2:${letter}`;
  };
  const name = `IF(${events("nameOverride")}<>"", ${events("nameOverride")}, ${events("defaultName")})`;
  return `=ARRAYFORMULA(IF(${id}2:${id}="", "", IFERROR(XLOOKUP(${id}2:${id}, ${
    events("fileId")
  }, ${name}), "(unknown event)")))`;
}
