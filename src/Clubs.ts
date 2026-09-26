/**
 * name of the Clubs tab
 */
const CLUBS_SHEET = "Clubs";

/**
 * text for the info row above the Clubs table
 * what the tab is, what to edit, how it refreshes
 */
const CLUBS_INFO = "Every known club and its teams, generated from Teams tab\n\n"
  + "DO NOT EDIT this table, it is refreshed by Refresh Tournaments, Refresh Results, and whenever a Club Override or Name Rule changes\n\n"
  + "Scores, ranks and the award are on the Club Statistics tab";
/**
 * clubs tab columns, in order, all plain values written by the script
 */
const CLUB_HEADERS = Object.freeze(
  {
    club: "Club",
    teams: "Teams",
    events: "Tournaments",
    teamCount: "Team Count",
  } as const,
);

/**
 * keys of the Clubs tab in column order
 */
const CLUB_KEYS = Object.keys(CLUB_HEADERS) as (keyof typeof CLUB_HEADERS)[];

/**
 * one row of the Clubs tab
 */
interface ClubRecord {
  /** club name */
  club: string;
  /** its teams, comma separated */
  teams: string;
  /** tournaments any of its teams played in, comma separated, in date order */
  events: string;
  /** how many teams */
  teamCount: number;
}

/**
 * @returns the Clubs tab, created on first use
 */
function _getClubsSheet_(): GoogleAppsScript.Spreadsheet.Sheet {
  const headers: readonly string[] = Object.values(CLUB_HEADERS);
  return _getOrCreateSheet_(SpreadsheetApp.getActiveSpreadsheet(), CLUBS_SHEET, headers, CLUBS_INFO);
}

/**
 * distinct club names, sorted A–Z
 * pure, never modifies its arguments
 * blanks ignored
 * matching is case-insensitive, first spelling seen is kept
 *
 * @param clubValues  values of the Teams Club column
 * @returns club names
 */
function _clubNames_(clubValues: unknown[]): string[] {
  const byKey: Map<string, string> = new Map();
  for (const value of clubValues) {
    const name = String(value).trim();
    const key = name.toLowerCase();
    if (name !== "" && !byKey.has(key)) byKey.set(key, name);
  }
  return [...byKey.values()].sort((a, b) => a.localeCompare(b, "en", { sensitivity: "base" }));
}

/**
 * every club with its teams and tournaments
 * pure, never modifies its arguments
 * club names are matched ignoring case, the first spelling seen is kept
 *
 * @param teamClubs  every team and its club, in display order
 * @param responses  all responses
 * @param events     all events, in date order
 * @returns one record per club, sorted A–Z
 */
function _clubRecords_(
  teamClubs: { team: string; club: string }[],
  responses: ResponseRecord[],
  events: EventRecord[],
): ClubRecord[] {
  const clubKeyOf = new Map(teamClubs.map((t) => [_teamKey_(t.team), t.club.toLowerCase()]));
  const played = _tournamentsPlayed_(responses, events, (team) => clubKeyOf.get(_teamKey_(team)) ?? "");
  return _clubNames_(teamClubs.map((t) => t.club)).map((club) => {
    const key = club.toLowerCase();
    const teams = teamClubs.filter((t) => t.club.toLowerCase() === key).map((t) => t.team);
    return { club, teams: teams.join(", "), events: played.get(key)?.names.join(", ") ?? "", teamCount: teams.length };
  });
}

/**
 * rebuild the Clubs tab from the Teams, Results and Tournaments tabs
 * one row per club, sorted A–Z, plain values so the text is easy to read
 * does not take the lock, callers are responsible
 * reads the Teams Club formula results, so call SpreadsheetApp.flush() first after writing Teams
 */
function _rebuildClubs_() {
  const events = _sortEvents_(_readEvents_(_getEventsSheet_()));
  const clubs = _clubRecords_(_readTeamClubList_(), _readResponses_(_getResponsesSheet_()), events);
  _writeTable_(_getClubsSheet_(), clubs.map((c) => CLUB_KEYS.map((key) => c[key])), CLUB_KEYS.length);
  _setClubReportChoices_(clubs.map((c) => c.club));
}
