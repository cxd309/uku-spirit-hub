/**
 * name of the Teams tab
 */
const TEAMS_SHEET = "Teams";

/**
 * text for the info row above the Teams table
 * what the tab is, what to edit, how it refreshes
 */
const TEAMS_INFO = "Every team that has given or recieved a spirit score (international clubs excluded)\n\n"
  + "Suggested Club name is generated using the enabled Name Rules on the Settings tab, this is the default \"best guess\" at a club\n\n"
  + "To override the club use the Club Override column\n\n"
  + "This table is refreshed by Refresh Tournaments and Refresh Results";

/**
 * teams tab columns, in order
 * columns listed in TEAM_FORMULAS are formulas
 * the rest hold TeamRecord values written by the script
 */
const TEAM_HEADERS = Object.freeze(
  {
    team: "Team",
    suggestedClub: "Suggested Club",
    clubOverride: "Club Override",
    club: "Club",
    events: "Tournaments",
    eventCount: "Event Count",
  } as const,
);

/**
 * keys of the Teams tab in column order
 */
const TEAM_KEYS = Object.keys(TEAM_HEADERS) as (keyof typeof TEAM_HEADERS)[];

/**
 * one team, as stored in js
 * its name and any club typed by a person
 */
interface TeamRecord {
  /** team name, as first seen */
  team: string;
  /** club typed by a person or "" if none */
  clubOverride: string;
  /** club suggested by the Name Rules, filled when written */
  suggestedClub?: string;
  /** tournaments played, comma separated, filled when written */
  events?: string;
  /** how many tournaments played, filled when written */
  eventCount?: number;
}

/**
 * formula columns of the Teams tab
 * Club is a formula so it follows a Club Override straight away
 * for each, a function building that column's formula for a given sheet row
 */
const TEAM_FORMULAS: Readonly<Partial<Record<keyof typeof TEAM_HEADERS, (row: number) => string>>> = Object.freeze({
  club: (row: number) => {
    const cell = (key: keyof typeof TEAM_HEADERS) => `$${_columnLetter_(TEAM_KEYS.indexOf(key) + 1)}${row}`;
    return `=IF(${cell("clubOverride")}<>"", ${cell("clubOverride")}, ${cell("suggestedClub")})`;
  },
});

/**
 * @returns the Teams tab, created on first use
 */
function _getTeamsSheet_(): GoogleAppsScript.Spreadsheet.Sheet {
  const headers: readonly string[] = Object.values(TEAM_HEADERS);
  return _getOrCreateSheet_(SpreadsheetApp.getActiveSpreadsheet(), TEAMS_SHEET, headers, TEAMS_INFO);
}

/**
 * read the stored columns (Team, Club Override) of every row on the Teams tab
 *
 * @param sheet  the Teams tab
 * @returns every team on the tab
 */
function _readTeams_(sheet: GoogleAppsScript.Spreadsheet.Sheet): TeamRecord[] {
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
 * @param team  the record
 * @param row   1-based sheet row it will be written to
 * @returns values in TEAM_KEYS order
 */
function _teamToRow_(team: TeamRecord, row: number): unknown[] {
  return TEAM_KEYS.map((key) => {
    const formula = TEAM_FORMULAS[key];
    return formula ? formula(row) : team[key as keyof TeamRecord];
  });
}

/**
 * write the team list, resizing the tab to fit
 * works out each team's suggested club and tournaments on the way
 *
 * @param sheet      the Teams tab
 * @param teams      teams to write, in display order
 * @param responses  all responses
 * @param events     all events, in date order
 */
function _writeTeams_(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  teams: TeamRecord[],
  responses: ResponseRecord[],
  events: EventRecord[],
) {
  const { regexes } = _readNameRules_();
  const played = _tournamentsPlayed_(responses, events, _teamKey_);
  _writeTable_(
    sheet,
    teams.map((t, i) => {
      const tournaments = played.get(_teamKey_(t.team));
      return _teamToRow_({
        ...t,
        suggestedClub: _suggestClub_(t.team, regexes),
        events: tournaments?.names.join(", ") ?? "",
        eventCount: tournaments?.count ?? 0,
      }, i + DATA_ROW);
    }),
    TEAM_KEYS.length,
  );
}

/**
 * the tournaments each group of teams played in
 * pure, never modifies its arguments
 * a team plays in a tournament when it gives or receives a score there
 *
 * @param responses  all responses
 * @param events     all events, in date order
 * @param groupOf    team name → group key, e.g. the team or its club, "" to leave it out
 * @returns group → tournament names in date order, and how many
 */
function _tournamentsPlayed_(
  responses: ResponseRecord[],
  events: EventRecord[],
  groupOf: (team: string) => string,
): Map<string, { names: string[]; count: number }> {
  const idsByGroup: Map<string, Set<string>> = new Map();
  for (const r of responses) {
    for (const team of [r.scorer, r.receiver]) {
      const group = groupOf(team);
      if (group !== "") idsByGroup.set(group, (idsByGroup.get(group) ?? new Set()).add(r.fileId));
    }
  }
  return new Map([...idsByGroup].map(([group, ids]) => {
    const names = events.filter((e) => ids.has(e.fileId) && e.tournament !== "").map((e) => e.tournament);
    return [group, { names: [...new Set(names)], count: ids.size }];
  }));
}

/**
 * work out every Suggested Club again from the current Name Rules
 * only the Suggested Club column is written
 */
function _refreshSuggestedClubs_() {
  const sheet = _getTeamsSheet_();
  const rowCount = sheet.getLastRow() - HEADER_ROW;
  if (rowCount < 1) return;
  const { regexes } = _readNameRules_();
  const teams = sheet.getRange(DATA_ROW, TEAM_KEYS.indexOf("team") + 1, rowCount, 1).getValues();
  sheet
    .getRange(DATA_ROW, TEAM_KEYS.indexOf("suggestedClub") + 1, rowCount, 1)
    .setValues(teams.map(([team]) => [_suggestClub_(String(team), regexes)]));
}

/**
 * key used to match team names, case-insensitive
 *
 * @param name  tidied team name
 * @returns match key
 */
function _teamKey_(name: string): string {
  return name.toLowerCase();
}

/**
 * every distinct team appearing in the responses
 * both teams count at national events
 * only the receiving (UK) team counts at international events
 * first spelling seen is kept
 *
 * @param responses         all responses
 * @param internationalIds  file ids of international events
 * @returns team names in order of first appearance
 */
function _teamNamesFromResponses_(responses: ResponseRecord[], internationalIds: Set<string>): string[] {
  const byKey: Map<string, string> = new Map();
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
 * @param existing  current rows of the Teams tab
 * @param names     team names from the responses
 * @returns the new team list
 */
function _mergeTeams_(existing: TeamRecord[], names: string[]): TeamRecord[] {
  const overrideByKey = new Map(existing.map((t) => [_teamKey_(t.team), t.clubOverride]));
  const currentKeys = new Set(names.map(_teamKey_));

  const current = names.map((team) => ({ team: team, clubOverride: overrideByKey.get(_teamKey_(team)) ?? "" }));
  const keptForOverride = existing.filter((t) => t.clubOverride !== "" && !currentKeys.has(_teamKey_(t.team)));

  return [...current, ...keptForOverride].sort((a, b) => a.team.localeCompare(b.team, "en", { sensitivity: "base" }));
}

/**
 * every team and its club, read from the Teams tab in display order
 * reads formula results, so call SpreadsheetApp.flush() first after writing Teams
 *
 * @returns one entry per team
 */
function _readTeamClubList_(): { team: string; club: string }[] {
  const sheet = _getTeamsSheet_();
  const rowCount = sheet.getLastRow() - HEADER_ROW;
  if (rowCount < 1) return [];
  const teamIndex = TEAM_KEYS.indexOf("team");
  const clubIndex = TEAM_KEYS.indexOf("club");
  return sheet
    .getRange(DATA_ROW, 1, rowCount, TEAM_KEYS.length)
    .getValues()
    .map((row) => ({ team: String(row[teamIndex]).trim(), club: String(row[clubIndex]).trim() }))
    .filter((t) => t.team !== "");
}

/**
 * club for every team, read from the Club column on the Teams tab
 * reads formula results, so call SpreadsheetApp.flush() first after writing Teams
 *
 * @returns team key → club
 */
function _readTeamClubs_(): Map<string, string> {
  return new Map(_readTeamClubList_().map((t) => [_teamKey_(t.team), t.club]));
}
