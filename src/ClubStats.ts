/**
 * name of the Club Statistics tab
 */
const CLUB_STATS_SHEET = "Club Statistics";

/**
 * club statistics columns, in order
 * keys are the names used in code
 * values are the header text
 */
const CLUB_STATS_HEADERS = Object.freeze(
  {
    club: "Club",
    tournamentsEntered: "Tournaments Entered",
    teamEntries: "Team Entries",
    qualifies: "Qualifies",
    responses: "Responses",
    mean: "Mean",
    sd: "SD",
    median: "Median",
    min: "Min",
    max: "Max",
    lowShare: "% 6 or Below",
    ciLower: "CI Lower",
    ciUpper: "CI Upper",
    rank: "Rank",
    clubModelMean: "Club Model Mean",
    clubModelSe: "Club Model SE",
    clubModelRank: "Club Model Rank",
    teamModelMean: "Team Model Mean",
    teamModelSe: "Team Model SE",
    teamModelRank: "Team Model Rank",
    averageGiven: "Average Given",
    scorerEffect: "Scorer Effect",
  } as const,
);

/**
 * keys of the Club Statistics tab in column order
 */
const CLUB_STATS_KEYS = Object.keys(CLUB_STATS_HEADERS) as (keyof typeof CLUB_STATS_HEADERS)[];

/**
 * a total at or below this counts towards % 6 or Below
 */
const CLUB_STATS_LOW_SCORE = 6;

/**
 * one row of the Club Statistics tab
 * numbers are "" when there is nothing to show
 */
type ClubStatsRow = Record<keyof typeof CLUB_STATS_HEADERS, string | number | boolean>;

/**
 * one score counted for club statistics, with both clubs worked out
 */
interface CountedScore {
  /** file id of the tournament */
  tournament: string;
  /** scoring team */
  scorer: string;
  /** scoring club, "" when the team has no club (e.g. international teams) */
  scorerClub: string;
  /** receiving team */
  receiver: string;
  /** receiving club */
  receiverClub: string;
  /** total score */
  total: number;
  /** true when the clubs differ, only these count towards scores and models */
  counts: boolean;
}

/**
 * info row text, with when it was calculated and a summary of both models
 *
 * @param when     when the statistics were calculated
 * @param minimum  tournaments needed to qualify
 * @param models   each model by name
 * @returns the info text
 */
function _clubStatsInfo_(when: Date, minimum: number, models: Record<string, ModelResult>): string {
  const sd = (variance: number) => Math.sqrt(variance).toFixed(2);
  const summaries = Object.entries(models).map(([name, model]) =>
    model.ok
      ? `- ${name}: scorer SD ${sd(model.scorerVariance)}, scorer at tournament SD ${
        sd(model.scorerAtTournamentVariance)
      }, remaining SD ${sd(model.residualVariance)}`
      : `- ${name}: not fitted, ${model.reason}`
  );
  return `Spirit statistics for each club, calculated ${
    Utilities.formatDate(when, Session.getScriptTimeZone(), "d MMM yyyy HH:mm")
  }\n\n`
    + "Counted: scores at included tournaments, between different clubs\n"
    + `Qualifies: entered at least ${minimum} tournaments, ranks are among qualifying clubs\n`
    + "CI: 95% confidence interval of the mean\n"
    + "Models: mean adjusted for scorers who give higher or lower scores, as in the old R script, "
    + "with scorers grouped by club or by team\n"
    + "Scorer Effect: how much higher (+) or lower (-) than average this club scores others, from the club model\n\n"
    + `Model spread:\n${summaries.join("\n")}\n\n`
    + "DO NOT EDIT, run \"Refresh Club Statistics\" to recalculate";
}

/**
 * @returns the Club Statistics tab, created on first use
 */
function _getClubStatsSheet_(): GoogleAppsScript.Spreadsheet.Sheet {
  const headers: readonly string[] = Object.values(CLUB_STATS_HEADERS);
  const info = "Spirit statistics for each club\n\nRun \"Refresh Club Statistics\" to calculate";
  return _getOrCreateSheet_(SpreadsheetApp.getActiveSpreadsheet(), CLUB_STATS_SHEET, headers, info);
}

/**
 * every score at an included tournament, with both clubs worked out
 * pure, never modifies its arguments
 * scores between two teams of the same club are kept for tournament counts but do not count, the same as the award
 *
 * @param responses  all responses
 * @param events     all events
 * @param clubOf     team key → club, from the Teams tab
 * @returns the scores, those with no receiving club dropped
 */
function _countedScores_(
  responses: ResponseRecord[],
  events: EventRecord[],
  clubOf: Map<string, string>,
): CountedScore[] {
  const included = new Set(events.filter((e) => e.include).map((e) => e.fileId));
  return responses
    .filter((r) => included.has(r.fileId))
    .map((r) => ({
      tournament: r.fileId,
      scorer: r.scorer,
      scorerClub: clubOf.get(_teamKey_(r.scorer)) ?? "",
      receiver: r.receiver,
      receiverClub: clubOf.get(_teamKey_(r.receiver)) ?? "",
      total: _responseTotal_(r),
    }))
    .filter((s) => s.receiverClub !== "")
    .map((s) => ({ ...s, counts: s.scorerClub !== s.receiverClub }));
}

/**
 * summary numbers for a list of totals
 * pure, never modifies its arguments
 *
 * @param totals  the totals
 * @returns each "" when there are too few totals
 */
function _summarise_(
  totals: number[],
): {
  mean: number | "";
  sd: number | "";
  median: number | "";
  min: number | "";
  max: number | "";
  lowShare: number | "";
} {
  const n = totals.length;
  if (n === 0) return { mean: "", sd: "", median: "", min: "", max: "", lowShare: "" };
  const sorted = [...totals].sort((a, b) => a - b);
  const mean = totals.reduce((a, b) => a + b, 0) / n;
  const sd = n < 2 ? "" : Math.sqrt(totals.reduce((sum, t) => sum + (t - mean) ** 2, 0) / (n - 1));
  const middle = Math.floor(n / 2);
  const median = n % 2 === 1 ? sorted[middle] : (sorted[middle - 1] + sorted[middle]) / 2;
  const lowShare = totals.filter((t) => t <= CLUB_STATS_LOW_SCORE).length / n;
  return { mean, sd, median, min: sorted[0], max: sorted[n - 1], lowShare };
}

/**
 * rank of each qualifying club by a value, highest first, ties share a rank
 * pure, never modifies its arguments
 *
 * @param rows  every club's row
 * @param key   column to rank by
 * @returns club → rank, qualifying clubs with a value only
 */
function _rankClubs_(rows: ClubStatsRow[], key: keyof typeof CLUB_STATS_HEADERS): Map<string, number> {
  const ranked = rows.filter((row) => row.qualifies === true && typeof row[key] === "number");
  return new Map(ranked.map((row) => [
    String(row.club),
    1 + ranked.filter((other) => (other[key] as number) > (row[key] as number)).length,
  ]));
}

/**
 * every club's statistics
 * pure, no google calls, except the two model fits which are also pure
 *
 * @param clubs    every club, in display order
 * @param scores   every score at an included tournament
 * @param minimum  tournaments needed to qualify
 * @returns one row per club, and both models
 */
function _clubStatistics_(
  clubs: string[],
  scores: CountedScore[],
  minimum: number,
): { rows: ClubStatsRow[]; models: Record<string, ModelResult> } {
  const counted = scores.filter((s) => s.counts);
  const observations = (scorer: (s: CountedScore) => string) =>
    counted.map((s) => ({ club: s.receiverClub, scorer: scorer(s), tournament: s.tournament, total: s.total }));
  // a scorer with no club, e.g. an international team, is its own group
  const clubModel = _fitClubModel_(
    observations((s) => s.scorerClub !== "" ? `club:${s.scorerClub}` : `team:${_teamKey_(s.scorer)}`),
  );
  const teamModel = _fitClubModel_(observations((s) => `team:${_teamKey_(s.scorer)}`));

  const rows: ClubStatsRow[] = clubs.map((club) => {
    const received = counted.filter((s) => s.receiverClub === club);
    const given = counted.filter((s) => s.scorerClub === club);
    const played = scores.filter((s) => s.receiverClub === club || s.scorerClub === club);
    // one entry per team per tournament, whichever side of the score the team is on
    const teamEntries: Set<string> = new Set();
    for (const s of played) {
      if (s.receiverClub === club) teamEntries.add(`${s.tournament}|${_teamKey_(s.receiver)}`);
      if (s.scorerClub === club) teamEntries.add(`${s.tournament}|${_teamKey_(s.scorer)}`);
    }
    const tournamentsEntered = new Set(played.map((s) => s.tournament)).size;
    const clubFit = clubModel.ok ? clubModel.clubs.get(club) : undefined;
    const teamFit = teamModel.ok ? teamModel.clubs.get(club) : undefined;
    const scorerEffect = clubModel.ok ? clubModel.scorerEffects.get(`club:${club}`) : undefined;
    const givenSummary = _summarise_(given.map((s) => s.total));

    return {
      club,
      tournamentsEntered,
      teamEntries: teamEntries.size,
      qualifies: tournamentsEntered >= minimum,
      responses: received.length,
      ..._summarise_(received.map((s) => s.total)),
      ciLower: "",
      ciUpper: "",
      rank: "",
      clubModelMean: clubFit?.mean ?? "",
      clubModelSe: clubFit?.se ?? "",
      clubModelRank: "",
      teamModelMean: teamFit?.mean ?? "",
      teamModelSe: teamFit?.se ?? "",
      teamModelRank: "",
      averageGiven: givenSummary.mean,
      scorerEffect: scorerEffect ?? "",
    };
  });

  const ranks = _rankClubs_(rows, "mean");
  const clubModelRanks = _rankClubs_(rows, "clubModelMean");
  const teamModelRanks = _rankClubs_(rows, "teamModelMean");
  return {
    rows: rows.map((row) => ({
      ...row,
      rank: ranks.get(String(row.club)) ?? "",
      clubModelRank: clubModelRanks.get(String(row.club)) ?? "",
      teamModelRank: teamModelRanks.get(String(row.club)) ?? "",
    })),
    models: { "Club model": clubModel, "Team model": teamModel },
  };
}

/**
 * club rows in rank order, best first
 * clubs without a rank (not qualifying) go last, and ties keep their A–Z order
 * pure, returns a new array
 *
 * @param rows  every club's row, A–Z
 * @returns the rows in rank order
 */
function _sortByRank_(rows: ClubStatsRow[]): ClubStatsRow[] {
  const rankOf = (row: ClubStatsRow) => (typeof row.rank === "number" ? row.rank : Number.MAX_SAFE_INTEGER);
  return [...rows].sort((a, b) => rankOf(a) - rankOf(b));
}

/**
 * formula for a CI cell, from the Responses, Mean and SD cells on its row
 * uses the t distribution, the same as R's group.CI
 *
 * @param row   1-based sheet row
 * @param sign  +1 for the upper limit, -1 for the lower
 * @returns the formula
 */
function _clubStatsCiFormula_(row: number, sign: 1 | -1): string {
  const cell = (key: keyof typeof CLUB_STATS_HEADERS) => `$${_columnLetter_(CLUB_STATS_KEYS.indexOf(key) + 1)}${row}`;
  const n = cell("responses");
  return `=IF(${n}>1, ${cell("mean")} ${sign > 0 ? "+" : "-"} T.INV.2T(0.05, ${n}-1)*${cell("sd")}/SQRT(${n}), "")`;
}

/**
 * recalculate the Club Statistics tab from the spreadsheet, no files are read
 *
 * @returns a one-line summary
 */
function _refreshClubStatistics_(): string {
  const events = _readEvents_(_getEventsSheet_());
  const responses = _readResponses_(_getResponsesSheet_());
  if (responses.length === 0) return "No results found: run Refresh Results first";

  const clubOf = _readTeamClubs_();
  const minimum = _readAwardSettings_().minimumTournaments;

  const scores = _countedScores_(responses, events, clubOf);
  const countedTotal = scores.filter((s) => s.counts).length;
  const { rows: clubRows, models } = _timed_(
    "club statistics",
    () => _clubStatistics_(_clubNames_([...clubOf.values()]), scores, minimum),
  );
  const rows = _sortByRank_(clubRows);

  const sheet = _getClubStatsSheet_();
  _writeTable_(
    sheet,
    rows.map((row, i) =>
      CLUB_STATS_KEYS.map((key) => {
        if (key === "ciLower") return _clubStatsCiFormula_(i + DATA_ROW, -1);
        if (key === "ciUpper") return _clubStatsCiFormula_(i + DATA_ROW, 1);
        return row[key];
      })
    ),
    CLUB_STATS_KEYS.length,
  );
  // no frozen columns, so the info row can be merged across the table like every other tab
  sheet.setFrozenColumns(0);
  sheet
    .getRange(INFO_ROW, 1, 1, CLUB_STATS_KEYS.length)
    .merge()
    .setValue(_clubStatsInfo_(new Date(), minimum, models))
    .setWrap(true)
    .setVerticalAlignment("top");

  if (rows.length > 0) {
    // full values are stored, decimals are shown to 2 decimal places
    const formats: [keyof typeof CLUB_STATS_HEADERS, string][] = [
      ["mean", "0.00"],
      ["sd", "0.00"],
      ["median", "0.0"],
      ["lowShare", "0%"],
      ["ciLower", "0.00"],
      ["ciUpper", "0.00"],
      ["clubModelMean", "0.00"],
      ["clubModelSe", "0.00"],
      ["teamModelMean", "0.00"],
      ["teamModelSe", "0.00"],
      ["averageGiven", "0.00"],
      ["scorerEffect", "+0.00;-0.00;0.00"],
    ];
    for (const [key, format] of formats) {
      sheet.getRange(DATA_ROW, CLUB_STATS_KEYS.indexOf(key) + 1, rows.length, 1).setNumberFormat(format);
    }
  }

  const failed = Object.entries(models).filter(([, m]) => !m.ok).map(([name]) => name);
  return `Statistics for ${rows.length} club(s) from ${countedTotal} score(s)`
    + (failed.length > 0 ? `; ${failed.join(" and ")} could not be fitted (see tab info)` : "");
}
