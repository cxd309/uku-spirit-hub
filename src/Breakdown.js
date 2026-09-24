/**
 * exact header text of each breakdown column from SOTG form
 * keys are names used in code
 * values must match row 1 exactly (after trimming) in sheet
 */
const BREAKDOWN_HEADERS = Object.freeze({
  yourTeam: "Your Team Name",
  opponentTeam: "Opponent Team Name",
  rules: "Opposition Rules Knowledge and Use",
  fouls: "Opposition Fouls and Body Contact",
  fairMindedness: "Opposition Fair-Mindedness",
  attitude: "Opposition Positive Attitude and Self-Control",
  communication: "Opposition Communication",
  comments: "Comments",
});

/**
 * keys of the five spirit score columns in form order
 */
const SCORE_KEYS = Object.freeze(
  /** @type{const} */ (["rules", "fouls", "fairMindedness", "attitude", "communication"]),
);

/**
 * 0-based column index of each breakdown column, keyed like BREAKDOWN-HEADERS
 *
 * @typedef {Record<keyof typeof BREAKDOWN_HEADERS, number>} BreakdownColumns
 */

/**
 * result of checking one header row
 *
 * @typedef {{ok: true, columns: BreakdownColumns} | {ok: false, reason: string}} HeaderMatch
 */

/**
 * result of looking for the breakdown tab in a results file
 *
 * @typedef {{ok:true, sheet: GoogleAppsScript.Spreadsheet.Sheet, columns: BreakdownColumns}
 * | {ok:false, reason:string}} BreakdownLookup
 */

/**
 * check whether a header row is a breakdown header, and where each column is
 *
 * each header in BREAKDOWN_HEADERS must appear exatly (cells are trimmed, extra columns ignored)
 * the team columns must be in the form: "Your Team Name" -> "Opponent Team Name"
 *
 * @param {unknown[]}     headerRow values of row 1
 * @returns {HeaderMatch} column positions, or why the row is not a breakdown header
 */
function _matchHeader_(headerRow) {
  const cells = headerRow.map((cell) => String(cell).trim());
  const keys = /** @type {(keyof typeof BREAKDOWN_HEADERS)[]} */ (Object.keys(BREAKDOWN_HEADERS));

  /** @type {Partial<BreakdownColumns>} */
  const columns = {};
  for (const key of keys) {
    const index = cells.indexOf(BREAKDOWN_HEADERS[key]);
    if (index === -1) {
      return { ok: false, reason: `missing column "${BREAKDOWN_HEADERS[key]}"` };
    }
    columns[key] = index;
  }
  return { ok: true, columns: /** @type {BreakdownColumns} */ (columns) };
}

/**
 * find the single breakdown tab in a results file
 *
 * only row 1 of each tab is read
 * fails if no tab matches or if more than one does
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} book  opened results file
 * @returns {BreakdownLookup} breakdown tab and its columns, or why it could not be found
 */
function _findBreakdownSheet_(book) {
  /** @type {{sheet: GoogleAppsScript.Spreadsheet.Sheet, columns: BreakdownColumns}[]} */
  const matches = [];
  /** @type {string[]} */
  const swapped = [];

  for (const sheet of book.getSheets()) {
    const width = sheet.getLastColumn();
    if (width === 0) continue;

    const header = sheet.getRange(1, 1, 1, width).getValues()[0];
    const match = _matchHeader_(header);
    if (match.ok) {
      matches.push({ sheet: sheet, columns: match.columns });
    } else if (match.reason.startsWith("team columns")) {
      swapped.push(`tab "${sheet.getName()}": ${match.reason}`);
    }
  }

  if (matches.length === 1) return { ok: true, ...matches[0] };
  if (matches.length > 1) {
    return {
      ok: false,
      reason: `${matches.length} tabs have breakdown headers: ${
        matches.map((m) => `"${m.sheet.getName()}"`).join(", ")
      }`,
    };
  }
  if (swapped.length > 0) return { ok: false, reason: swapped.join("; ") };
  return { ok: false, reason: "no tab has the breakdown headers" };
}

/**
 * the five category scores for one game, each a whole number 0–4
 *
 * @typedef {Record<typeof SCORE_KEYS[number], number>} SpiritScores
 */

/**
 * one usable row from a breakdown tab
 *
 * @typedef {Object} SpiritResponse
 * @property {number}      sourceRow  1-based row number in the breakdown tab
 * @property {string}      scorer     team giving the score ("Your Team Name")
 * @property {string}      receiver   team receiving the score ("Opponent Team Name")
 * @property {SpiritScores} scores    the five category scores
 * @property {string}      comment    spirit comment, "" if blank
 */

/**
 * a row that could not be read, and why
 *
 * @typedef {Object} RowProblem
 * @property {number} sourceRow  1-based row number in the breakdown tab
 * @property {string} reason     what is wrong with the row
 */

/**
 * whether a cell holds a valid spirit score: a whole number from 0 to 4
 *
 * @param {unknown} value cell value
 * @return {boolean} true if valid
 */
function _isValidScore_(value) {
  return typeof value === "number" && Number.isInteger(value) && value >= 0 && value <= 4;
}

/**
 * turn the values of a breakdown tab into responses
 *
 * row 1 is the header and is skipped
 * completely blank rows are skipped
 * rows with a missing team name or an invalid score are returned as problems instead of responses
 *
 * @param {unknown[][]}      values   all values of the breakdown tab, including the header row
 * @param {BreakdownColumns} columns  column positions from _matchHeader_
 * @returns {{responses: SpiritResponse[], problems: RowProblem[]}} usable rows and unusable rows
 */
function _parseBreakdownRows_(values, columns) {
  /** @type {SpiritResponse[]} */
  const responses = [];
  /** @type {RowProblem[]} */
  const problems = [];

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const sourceRow = i + 1;
    const scorer = String(row[columns.yourTeam]).trim();
    const receiver = String(row[columns.opponentTeam]).trim();
    const comment = String(row[columns.comments]).trim();

    const isBlank = !scorer && !receiver && !comment && SCORE_KEYS.every((k) => row[columns[k]] === "");
    if (isBlank) continue;

    /** @type {string[]} */
    const reasons = [];
    if (!scorer) reasons.push(`missing ${BREAKDOWN_HEADERS.yourTeam}`);
    if (!receiver) reasons.push(`missing ${BREAKDOWN_HEADERS.opponentTeam}`);
    const badScores = SCORE_KEYS.filter((k) => !_isValidScore_(row[columns[k]]));
    if (badScores.length > 0) {
      reasons.push(
        `invalid score in ${
          badScores.map((k) => `"${BREAKDOWN_HEADERS[k]}" (${JSON.stringify(row[columns[k]])})`).join(", ")
        }`,
      );
    }

    if (reasons.length > 0) {
      problems.push({ sourceRow: sourceRow, reason: reasons.join("; ") });
      continue;
    }

    responses.push({
      sourceRow: sourceRow,
      scorer: scorer,
      receiver: receiver,
      scores: {
        rules: Number(row[columns.rules]),
        fouls: Number(row[columns.fouls]),
        fairMindedness: Number(row[columns.fairMindedness]),
        attitude: Number(row[columns.attitude]),
        communication: Number(row[columns.communication]),
      },
      comment: comment,
    });
  }
  return { responses: responses, problems: problems };
}
