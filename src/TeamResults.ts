/**
 * name of the Team Results tab
 */
const TEAM_RESULTS_SHEET = "Team Results";

/**
 * team results columns, in order
 * keys are the names used in code
 * values are the header text
 */
const TEAM_RESULTS_HEADERS = Object.freeze(
  {
    tournament: "Tournament",
    date: "Date",
    division: "Div",
    team: "Team",
    club: "Club",
    scores: "Scores",
    mean: "Mean",
    median: "Median",
    rules: "Rules",
    fouls: "Fouls",
    fairMindedness: "Fair-Mindedness",
    attitude: "Attitude",
    communication: "Communication",
    position: "Position",
    reversePosition: "Reverse Position",
    teams: "Teams",
  } as const,
);

/**
 * keys of the Team Results tab in column order
 */
const TEAM_RESULTS_KEYS = Object.keys(TEAM_RESULTS_HEADERS) as (keyof typeof TEAM_RESULTS_HEADERS)[];

/**
 * one row of the Team Results tab: one receiving team at one tournament
 */
type TeamResultRow = Record<keyof typeof TEAM_RESULTS_HEADERS, string | number | Date>;

/**
 * info row text, with when it was calculated
 *
 * @param when  when the table was calculated
 * @returns the info text
 */
function _teamResultsInfo_(when: Date): string {
  return `Each team's spirit result at each tournament, calculated ${
    Utilities.formatDate(when, Session.getScriptTimeZone(), "d MMM yyyy HH:mm")
  }\n\n`
    + "Counts the same scores as the published Results file: forms ticked Counted on Results, "
    + "at tournaments ticked Include Results, same club scores included\n"
    + "Mean, Median and the categories are of the scores the team received\n"
    + "Position: 1 is the highest Mean at the tournament, Reverse Position: 1 is the lowest, "
    + "ties share a position, Teams: teams with a score at the tournament\n\n"
    + "DO NOT EDIT, run \"Refresh Club Statistics\" to recalculate";
}

/**
 * @returns the Team Results tab, created on first use
 */
function _getTeamResultsSheet_(): GoogleAppsScript.Spreadsheet.Sheet {
  const headers: readonly string[] = Object.values(TEAM_RESULTS_HEADERS);
  const info = "Each team's spirit result at each tournament\n\nRun \"Refresh Club Statistics\" to calculate";
  return _getOrCreateSheet_(SpreadsheetApp.getActiveSpreadsheet(), TEAM_RESULTS_SHEET, headers, info);
}

/**
 * every receiving team's result at every included tournament
 * pure, never modifies its arguments
 * the scores are those of the published Results file: counted forms, at tournaments ticked Include Results
 *
 * @param events     all events, in date order
 * @param responses  all responses
 * @param clubOf     team key → club, from the Teams tab
 * @returns rows by tournament in date order, then by position
 */
function _teamResults_(
  events: EventRecord[],
  responses: ResponseRecord[],
  clubOf: Map<string, string>,
): TeamResultRow[] {
  const counted = responses.filter((r) => r.counted);
  return events.filter((e) => e.includeResults).flatMap((event) => {
    const byTeam: Map<string, ResponseRecord[]> = new Map();
    for (const r of counted) {
      if (r.spiritFormUrl === event.spiritFormUrl) byTeam.set(r.receiver, [...(byTeam.get(r.receiver) ?? []), r]);
    }
    const teams = [...byTeam].map(([team, forms]) => {
      const summary = _summarise_(forms.map(_responseTotal_));
      const average = (key: (typeof SCORE_KEYS)[number]) => forms.reduce((sum, r) => sum + r[key], 0) / forms.length;
      return {
        team,
        forms,
        mean: summary.mean as number,
        median: summary.median as number,
        categories: SCORE_KEYS.map(average),
      };
    });
    return teams
      .map((t) => ({
        tournament: event.eventName,
        date: event.startDate ?? "",
        division: event.division,
        team: t.team,
        club: clubOf.get(_teamKey_(t.team)) ?? "",
        scores: t.forms.length,
        mean: t.mean,
        median: t.median,
        rules: t.categories[0],
        fouls: t.categories[1],
        fairMindedness: t.categories[2],
        attitude: t.categories[3],
        communication: t.categories[4],
        position: 1 + teams.filter((o) => o.mean > t.mean).length,
        reversePosition: 1 + teams.filter((o) => o.mean < t.mean).length,
        teams: teams.length,
      }))
      .sort((a, b) => a.position - b.position || a.team.localeCompare(b.team));
  });
}

/**
 * rebuild the Team Results tab
 * written as values, so it only changes when Refresh Club Statistics runs
 *
 * @param events     all events, in date order
 * @param responses  all responses
 * @param clubOf     team key → club, from the Teams tab
 */
function _refreshTeamResults_(events: EventRecord[], responses: ResponseRecord[], clubOf: Map<string, string>) {
  const rows = _teamResults_(events, responses, clubOf);
  const sheet = _getTeamResultsSheet_();
  const columns = TEAM_RESULTS_KEYS.length;
  _writeTable_(sheet, rows.map((row) => TEAM_RESULTS_KEYS.map((key) => row[key])), columns);
  sheet
    .getRange(INFO_ROW, 1, 1, columns)
    .merge()
    .setValue(_teamResultsInfo_(new Date()))
    .setWrap(true)
    .setVerticalAlignment("top");

  if (rows.length > 0) {
    // full values are stored, decimals are shown to 2 decimal places
    const formats: [keyof typeof TEAM_RESULTS_HEADERS, string][] = [
      ["date", "yyyy-mm-dd"],
      ["mean", "0.00"],
      ["median", "0.0"],
      ...SCORE_KEYS.map((key) => [key, "0.00"] as [keyof typeof TEAM_RESULTS_HEADERS, string]),
    ];
    for (const [key, format] of formats) {
      sheet.getRange(DATA_ROW, TEAM_RESULTS_KEYS.indexOf(key) + 1, rows.length, 1).setNumberFormat(format);
    }
  }
}
