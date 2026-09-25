/**
 * name of the Clubs tab
 */
const CLUBS_SHEET = "Clubs";

/**
 * text for the info row above the Clubs table
 * what the tab is, what to edit, how it refreshes
 */
const CLUBS_INFO = "Every known club and its teams, generated from Teams tab\n\n"
  + "DO NOT EDIT this table, it is refreshed every time there is a change in the Teams tab";
/**
 * clubs tab columns, in order
 * columns listed in CLUB_FORMULAS are formulas
 * the rest hold ClubRecord values
 */
const CLUB_HEADERS = Object.freeze(
  /** @type {const} */ ({
    club: "Club",
    teams: "Teams",
    events: "Tournaments",
    teamCount: "Team Count",
    tournamentsEntered: "Tournaments Entered",
    teamEntries: "Team Entries",
  }),
);

/**
 * keys of the Clubs tab in column order
 */
const CLUB_KEYS = /** @type {(keyof typeof CLUB_HEADERS)[]} */ (Object.keys(CLUB_HEADERS));

/**
 * one club, as stored in js
 *
 * @typedef {Object} ClubRecord
 * @property {string} club  club name
 */

/**
 * references shared by the Clubs formulas for one row
 *
 * @typedef {Object} ClubFormulaRefs
 * @property {string}                                        club       this row's club cell, e.g. $A5
 * @property {function(keyof typeof TEAM_HEADERS): string}     teams      a Teams column below its header
 * @property {function(keyof typeof RESPONSE_HEADERS): string} responses  a Responses column below its header
 */

/**
 * build the references used by the Clubs formulas for one row
 *
 * @param {number} row  1-based sheet row
 * @returns {ClubFormulaRefs} references for that row
 */
function _clubFormulaRefs_(row) {
  const club = `$${_columnLetter_(CLUB_KEYS.indexOf("club") + 1)}${row}`;
  /** @param {keyof typeof TEAM_HEADERS} key */
  const teams = (key) => _columnBelowHeader_(TEAMS_SHEET, TEAM_KEYS.indexOf(key) + 1);
  /** @param {keyof typeof RESPONSE_HEADERS} key */
  const responses = (key) => _columnBelowHeader_(RESPONSES_SHEET, RESPONSE_KEYS.indexOf(key) + 1);
  return { club, teams, responses };
}

/**
 * formula columns of the Clubs tab
 * for each, a function building that column's formula for a given sheet row
 * tournaments entered counts distinct file ids where either side is the club
 * team entries sums the event count of the club's teams
 *
 * @type {Readonly<Partial<Record<keyof typeof CLUB_HEADERS, function(number): string>>>}
 */
const CLUB_FORMULAS = Object.freeze({
  teams: (/** @type {number} */ row) => {
    const { club, teams } = _clubFormulaRefs_(row);
    return `=IFERROR(TEXTJOIN(", ", TRUE, FILTER(${teams("team")}, ${teams("club")}=${club})), "")`;
  },
  teamCount: (/** @type {number} */ row) => {
    const { club, teams } = _clubFormulaRefs_(row);
    return `=COUNTIF(${teams("club")}, ${club})`;
  },
  tournamentsEntered: (/** @type {number} */ row) => {
    const { club, responses } = _clubFormulaRefs_(row);
    return `=IFERROR(COUNTUNIQUE(FILTER(${responses("fileId")}, `
      + `(${responses("scorerClub")}=${club})+(${responses("receiverClub")}=${club}))), 0)`;
  },
  teamEntries: (/** @type {number} */ row) => {
    const { club, teams } = _clubFormulaRefs_(row);
    return `=SUMIF(${teams("club")}, ${club}, ${teams("eventCount")})`;
  },
  events: (/** @type {number} */ row) => {
    const { club, responses } = _clubFormulaRefs_(row);
    return `=IFERROR(TEXTJOIN(", ", TRUE, UNIQUE(FILTER(${responses("tournament")}, `
      + `(${responses("scorerClub")}=${club})+(${responses("receiverClub")}=${club})))), "")`;
  },
});

/**
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} the Clubs tab, created on first use
 */
function _getClubsSheet_() {
  /** @type {readonly string[]} */
  const headers = Object.values(CLUB_HEADERS);
  return _getOrCreateSheet_(SpreadsheetApp.getActiveSpreadsheet(), CLUBS_SHEET, headers, CLUBS_INFO).sheet;
}

/**
 * distinct club names, sorted A–Z
 * pure, never modifies its arguments
 * blanks ignored
 * matching is case-insensitive, first spelling seen is kept
 *
 * @param {unknown[]} clubValues  values of the Teams Club column
 * @returns {string[]} club names
 */
function _clubNames_(clubValues) {
  /** @type {Map<string, string>} */
  const byKey = new Map();
  for (const value of clubValues) {
    const name = String(value).trim();
    const key = name.toLowerCase();
    if (name !== "" && !byKey.has(key)) byKey.set(key, name);
  }
  return [...byKey.values()].sort((a, b) => a.localeCompare(b, "en", { sensitivity: "base" }));
}

/**
 * convert a ClubRecord into a row of values, with formulas in formula columns
 *
 * @param {ClubRecord} club  the record
 * @param {number}     row   1-based sheet row it will be written to
 * @returns {unknown[]} values in CLUB_KEYS order
 */
function _clubToRow_(club, row) {
  return CLUB_KEYS.map((key) => {
    const formula = CLUB_FORMULAS[key];
    return formula ? formula(row) : club[/** @type {keyof ClubRecord} */ (key)];
  });
}

/**
 * rebuild the Clubs tab from the current Club values on the Teams tab
 * one row per club, sorted A–Z, each with its own formulas
 * does not take the lock, callers are responsible
 */
function _rebuildClubs_() {
  const teamsSheet = _getTeamsSheet_();
  const teamRows = teamsSheet.getLastRow() - HEADER_ROW;
  const clubColumn = TEAM_KEYS.indexOf("club") + 1;
  const clubValues = teamRows < 1
    ? []
    : teamsSheet.getRange(DATA_ROW, clubColumn, teamRows, 1).getValues().map((r) => r[0]);

  const clubs = _clubNames_(clubValues).map((name) => ({ club: name }));
  _writeTable_(_getClubsSheet_(), clubs.map((c, i) => _clubToRow_(c, i + DATA_ROW)), CLUB_KEYS.length);
}
