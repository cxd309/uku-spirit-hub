/**
 * name of the Teams tab
 */
const TEAMS_SHEET = "Teams";

/**
 * text for the info row above the Teams table
 * what the tab is, what to edit, how it refreshes
 */
const TEAMS_INFO = "Every team that has given or recieved a spirit score (international clubs excluded)\n\n"
  + "Suggested Club name is generated using the selected rules in the Name Rules tab this is the default \"best guess\" at a club\n\n"
  + "To override the club use the Club Override column\n\n"
  + "This table will be refreshed with the Responses table";

/**
 * teams tab columns, in order
 * columns listed in TEAM_FORMULAS are formulas
 * the rest hold TeamRecord values
 */
const TEAM_HEADERS = Object.freeze(
  /** @type {const} */ ({
    team: "Team",
    suggestedClub: "Suggested Club",
    clubOverride: "Club Override",
    club: "Club",
    events: "Tournaments",
    eventCount: "Event Count",
  }),
);

/**
 * keys of the Teams tab in column order
 */
const TEAM_KEYS = /** @type {(keyof typeof TEAM_HEADERS)[]} */ (Object.keys(TEAM_HEADERS));

/**
 * one team, as stored in js
 * its name and any club typed by a person
 *
 * @typedef {Object} TeamRecord
 * @property {string} team          team name, as first seen
 * @property {string} clubOverride  club typed by a person or "" if none
 */

/**
 * references shared by the Teams formulas for one row
 *
 * @typedef {Object} TeamFormulaRefs
 * @property {string}                                           team       this row's team cell, e.g. $A5
 * @property {function(keyof typeof TEAM_HEADERS): string}      cell       another cell on this row
 * @property {function(keyof typeof RESPONSE_HEADERS): string}  responses  a whole Responses column
 * @property {string}                                           plays      condition that is non-zero for Responses rows where this team scored or was scored
 */

/**
 * build the references used by the Teams formulas for one row
 *
 * @param {number} row  1-based sheet row
 * @returns {TeamFormulaRefs} references for that row
 */
function _teamFormulaRefs_(row) {
  /** @param {keyof typeof TEAM_HEADERS} key */
  const cell = (key) => `$${_columnLetter_(TEAM_KEYS.indexOf(key) + 1)}${row}`;
  /** @param {keyof typeof RESPONSE_HEADERS} key */
  const responses = (key) => _columnBelowHeader_(RESPONSES_SHEET, RESPONSE_KEYS.indexOf(key) + 1);
  const team = cell("team");
  const plays = `(${responses("scorer")}=${team})+(${responses("receiver")}=${team})`;
  return { team, cell, responses, plays };
}

/**
 * formula columns of the Teams tab
 * for each, a function building that column's formula for a given sheet row
 *
 * @type {Readonly<Partial<Record<keyof typeof TEAM_HEADERS, function(number): string>>>}
 */
const TEAM_FORMULAS = Object.freeze({
  events: (/** @type {number} */ row) => {
    const { responses, plays } = _teamFormulaRefs_(row);
    return `=IFERROR(TEXTJOIN(", ", TRUE, UNIQUE(FILTER(${responses("tournament")}, ${plays}))), "")`;
  },
  eventCount: (/** @type {number} */ row) => {
    const { responses, plays } = _teamFormulaRefs_(row);
    return `=IFERROR(COUNTUNIQUE(FILTER(${responses("fileId")}, ${plays})), 0)`;
  },
  suggestedClub: (/** @type {number} */ row) => {
    const { team } = _teamFormulaRefs_(row);
    const patterns = _nameRuleColumn_("Pattern");
    const enabled = _nameRuleColumn_("Enabled");
    return `=IFERROR(REDUCE(${team}, FILTER(${patterns}, ${enabled}=TRUE), `
      + `LAMBDA(name, pattern, TRIM(REGEXREPLACE(name, pattern, "")))), ${team})`;
  },
  club: (/** @type {number} */ row) => {
    const { cell } = _teamFormulaRefs_(row);
    return `=IF(${cell("clubOverride")}<>"", ${cell("clubOverride")}, ${cell("suggestedClub")})`;
  },
});

/**
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} the Teams tab, created on first use
 */
function _getTeamsSheet_() {
  /** @type {readonly string[]} */
  const headers = Object.values(TEAM_HEADERS);
  return _getOrCreateSheet_(SpreadsheetApp.getActiveSpreadsheet(), TEAMS_SHEET, headers, TEAMS_INFO).sheet;
}

/**
 * read the stored columns (Team, Club Override) of every row on the Teams tab
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet  the Teams tab
 * @returns {TeamRecord[]} every team on the tab
 */
function _readTeams_(sheet) {
  const rowCount = sheet.getLastRow() - HEADER_ROW;
  if (rowCount < 1) return [];
  const teamIndex = TEAM_KEYS.indexOf("team");
  const overrideIndex = TEAM_KEYS.indexOf("clubOverride");
  return sheet
    .getRange(DATA_ROW, 1, rowCount, TEAM_KEYS.length)
    .getValues()
    .map((row) => ({ team: String(row[teamIndex]).trim(), clubOverride: String(row[overrideIndex]).trim() }))
    .filter((t) => t.team !== "");
}

/**
 * convert a TeamRecord into a row of values, with formulas in formula columns
 *
 * @param {TeamRecord} team  the record
 * @param {number}     row   1-based sheet row it will be written to
 * @returns {unknown[]} balues in TEAM_KEYS order
 */
function _teamToRow_(team, row) {
  return TEAM_KEYS.map((key) => {
    const formula = TEAM_FORMULAS[key];
    return formula ? formula(row) : team[/** @type {keyof TeamRecord} */ (key)];
  });
}

/**
 * write the team list, resizing the tab to fit
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet  the Teams tab
 * @param {TeamRecord[]}                       teams  teams to write, in display order
 */
function _writeTeams_(sheet, teams) {
  _writeTable_(sheet, teams.map((t, i) => _teamToRow_(t, i + DATA_ROW)), TEAM_KEYS.length);
}

/**
 * key used to match team names, case-insensitive
 *
 * @param {string} name  tidied team name
 * @returns {string} match key
 */
function _teamKey_(name) {
  return name.toLowerCase();
}

/**
 * every distinct team appearing in the responses
 * both teams count at national events
 * only the receiving (UK) team counts at international events
 * first spelling seen is kept
 *
 * @param {ResponseRecord[]} responses         all responses
 * @param {Set<string>}      internationalIds  file ids of international events
 * @returns {string[]} team names in order of first appearance
 */
function _teamNamesFromResponses_(responses, internationalIds) {
  /** @type {Map<string, string>} */
  const byKey = new Map();
  for (const r of responses) {
    const names = internationalIds.has(r.fileId) ? [r.receiver] : [r.scorer, r.receiver];
    for (const name of names) {
      const key = _teamKey_(name);
      if (!byKey.has(key)) byKey.set(key, name);
    }
  }
  return [...byKey.values()];
}

/**
 * combine the current team names with the existing Teams tab
 * pure, never modifies its arguments
 * every current team is kept with its club override
 * teams no longer in any response are dropped unless they have a club override
 * sorted A–Z
 *
 * @param {TeamRecord[]} existing  current rows of the Teams tab
 * @param {string[]}     names     team names from the responses
 * @returns {TeamRecord[]} the new team list
 */
function _mergeTeams_(existing, names) {
  const overrideByKey = new Map(existing.map((t) => [_teamKey_(t.team), t.clubOverride]));
  const currentKeys = new Set(names.map(_teamKey_));

  const current = names.map((team) => ({ team: team, clubOverride: overrideByKey.get(_teamKey_(team)) ?? "" }));
  const keptForOverride = existing.filter((t) => t.clubOverride !== "" && !currentKeys.has(_teamKey_(t.team)));

  return [...current, ...keptForOverride].sort((a, b) => a.team.localeCompare(b.team, "en", { sensitivity: "base" }));
}

/**
 * club for every team, read from the Club column on the Teams tab
 * reads formula results, so call SpreadsheetApp.flush() first after writing Teams
 *
 * @returns {Map<string, string>} team key → club
 */
function _readTeamClubs_() {
  const sheet = _getTeamsSheet_();
  const rowCount = sheet.getLastRow() - HEADER_ROW;
  if (rowCount < 1) return new Map();
  const teamIndex = TEAM_KEYS.indexOf("team");
  const clubIndex = TEAM_KEYS.indexOf("club");
  return new Map(
    sheet.getRange(DATA_ROW, 1, rowCount, TEAM_KEYS.length).getValues()
      .map((row) => [_teamKey_(String(row[teamIndex])), String(row[clubIndex])]),
  );
}
