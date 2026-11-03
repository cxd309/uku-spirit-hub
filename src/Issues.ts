/**
 * name of the Issues tab
 */
const ISSUES_SHEET = "Issues";

/**
 * text for the info row above the Issues table
 * what the tab is, what to edit, how it refreshes
 */
const ISSUES_INFO = "Spirit issues found in the results\n"
  + "Play: one team in one game, from the scores it received. Scoring: one team in one game, from the form it gave. "
  + "Tournament: one team at one tournament\n"
  + "Game is blank and Opponent is used instead when a form has no game, e.g. a duplicate or a deleted form\n\n"
  + "User editable columns:\n"
  + "- Status: NEW, IN PROGRESS or CLOSED. Change it once you have sent the email\n"
  + "- Committee Member: who is handling it\n"
  + "- Notes\n\n"
  + "A refresh that finds another check for an existing issue adds it to Checks and Details, "
  + "and a CLOSED issue goes back to NEW. Nothing else on existing rows is changed, "
  + "apart from the Open in mail link for NEW issues whose email template is enabled\n\n"
  + "New issues are added by Refresh Issues, so sort and filter freely";

/**
 * issues table columns, in order
 * keys are the names used in code
 * values are the header text
 */
const ISSUE_HEADERS = Object.freeze(
  {
    issueId: "Issue ID",
    dateCreated: "Issue Date",
    level: "Level",
    checks: "Checks",
    tournament: "Tournament",
    tournamentDate: "Tournament Date",
    game: "Game",
    opponent: "Opponent",
    team: "Team",
    club: "Club",
    details: "Details",
    status: "Status",
    email: "Email",
    committeeMember: "Committee Member",
    notes: "Notes",
  } as const,
);

/**
 * keys of the Issues tab in column order
 */
const ISSUE_KEYS = Object.keys(ISSUE_HEADERS) as (keyof typeof ISSUE_HEADERS)[];

/**
 * 0-based columns a person edits, left white, every other column is shaded
 */
const ISSUE_EDITABLE = Object.freeze(
  (["status", "committeeMember", "notes"] as const).map((key) => ISSUE_KEYS.indexOf(key)),
);

/**
 * values of the Status column, first is given to new issues
 */
const ISSUE_STATUSES = Object.freeze(["NEW", "IN PROGRESS", "CLOSED"]);

/**
 * values of the Level column
 * Play and Scoring issues are about one team in one game, Tournament issues one team at one tournament
 */
const ISSUE_LEVELS = Object.freeze({ PLAY: "Play", SCORING: "Scoring", TOURNAMENT: "Tournament" } as const);

/**
 * every issue check
 * code goes into the Checks column, the Issue ID of tournament issues and the Rule column of the Issue Rules settings
 * level says which kind of issue it adds to
 * label opens its section of Details, built from the current settings
 */
const ISSUE_CATEGORIES = Object.freeze({
  lowScore: {
    code: "LOW-SCORE",
    level: ISSUE_LEVELS.PLAY,
    label: (s: IssueSettings) => `Received a total of ${s.lowScoreAtOrBelow} or below`,
  },
  dangerousPlay: {
    code: "DANGEROUS-PLAY",
    level: ISSUE_LEVELS.PLAY,
    label: (s: IssueSettings) => "Comment mentions dangerous play",
  },
  cheating: {
    code: "CHEATING",
    level: ISSUE_LEVELS.PLAY,
    label: (s: IssueSettings) => "Comment mentions cheating",
  },
  harassment: {
    code: "HARASSMENT",
    level: ISSUE_LEVELS.PLAY,
    label: (s: IssueSettings) => "Comment mentions harassment",
  },
  categoryMinimum: {
    code: "MIN-CATEGORY",
    level: ISSUE_LEVELS.PLAY,
    label: (s: IssueSettings) => `Received ${s.categoryMinimum} in a category`,
  },
  totalWithoutComment: {
    code: "TOTAL-NO-COMMENT",
    level: ISSUE_LEVELS.SCORING,
    label: (s: IssueSettings) =>
      `Gave a total above ${s.commentTotalAbove} or below ${s.commentTotalBelow} without comment`,
  },
  categoryWithoutComment: {
    code: "CATEGORY-NO-COMMENT",
    level: ISSUE_LEVELS.SCORING,
    label: (s: IssueSettings) => `Gave ${s.commentCategoryScores.join(" or ")} in a category without comment`,
  },
  deletedRow: {
    code: "DELETED-ROW",
    level: ISSUE_LEVELS.SCORING,
    label: (s: IssueSettings) => "Form deleted",
  },
  selfScore: {
    code: "SELF-SCORE",
    level: ISSUE_LEVELS.SCORING,
    label: (s: IssueSettings) => "Scored itself",
  },
  noSuchGame: {
    code: "NO-SUCH-GAME",
    level: ISSUE_LEVELS.SCORING,
    label: (s: IssueSettings) => "Form for a game that is not on the schedule",
  },
  duplicate: {
    code: "DUPLICATE",
    level: ISSUE_LEVELS.SCORING,
    label: (s: IssueSettings) => "More forms than games",
  },
  noScores: {
    code: "NO-SCORES",
    level: ISSUE_LEVELS.TOURNAMENT,
    label: (s: IssueSettings) => "No spirit scores submitted",
  },
  underSubmitted: {
    code: "UNDER-SUBMITTED",
    level: ISSUE_LEVELS.TOURNAMENT,
    label: (s: IssueSettings) => `Spirit scores submitted for fewer than ${s.submittedPercentBelow}% of games`,
  },
  immediateAverage: {
    code: "IMMEDIATE-AVERAGE",
    level: ISSUE_LEVELS.TOURNAMENT,
    label: (s: IssueSettings) => `Average below ${s.immediateAverageBelow}`,
  },
  monitoringAverage: {
    code: "MONITORING-AVERAGE",
    level: ISSUE_LEVELS.TOURNAMENT,
    label: (s: IssueSettings) => `Average below ${s.monitoringAverageBelow}`,
  },
});

/**
 * checks raised by words in a comment, each with its own word list in the Issue Rules settings
 * the issue is for the receiving team
 */
const COMMENT_KEYWORD_CHECKS = Object.freeze(["dangerousPlay", "cheating", "harassment"] as const);

/**
 * a pattern matching any of the words, as whole words, ignoring case
 * pure, no google calls
 * works for words ending in punctuation too
 *
 * @param words  the words
 * @returns the pattern, null when there are no words
 */
function _keywordPattern_(words: string[]): RegExp | null {
  if (words.length === 0) return null;
  const escaped = words.map((word) => word.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"));
  // whole words only, not part of a longer word
  return new RegExp(`(?<!\\w)(${escaped.join("|")})(?!\\w)`, "i");
}

/**
 * key of one issue check, e.g. "lowScore"
 */
type IssueCategoryKey = keyof typeof ISSUE_CATEGORIES;

/**
 * one row of the Issues tab
 */
interface IssueRecord {
  /** level, tournament, game or opponent, and team, unique per issue */
  issueId: string;
  /** when the issue was first added */
  dateCreated: Date;
  /** one of ISSUE_LEVELS */
  level: string;
  /** codes of every check found, comma separated */
  checks: string;
  /** tournament name */
  tournament: string;
  /** tournament start date, null if unknown */
  tournamentDate: Date | null;
  /** game code, "" for tournament issues and forms with no game */
  game: string;
  /** the other team in the game, "" for tournament issues */
  opponent: string;
  /** team the issue is about */
  team: string;
  /** club of the team */
  club: string;
  /** one section per check, then the comments */
  details: string;
  /** one of ISSUE_STATUSES */
  status: string;
  /** Open in mail link, filled in later by _updateEmailLinks_ */
  email: string;
  /** who is handling it */
  committeeMember: string;
  /** free text */
  notes: string;
}

/**
 * one thing a check found, before findings are grouped into issues
 */
interface IssueFinding {
  /** which check */
  category: IssueCategoryKey;
  /** the tournament */
  event: EventRecord;
  /** team the issue is about */
  team: string;
  /** game code, "" when there is none */
  game: string;
  /** the other team, "" for tournament checks */
  opponent: string;
  /** what was found, for Details */
  text: string;
  /** comment to quote below the sections, with who wrote it, null for none */
  comment: { from: string; text: string } | null;
}

/**
 * total of the five category scores in a response
 *
 * @param response  the response
 * @returns the total
 */
function _responseTotal_(response: ResponseRecord): number {
  return SCORE_KEYS.reduce((sum, key) => sum + response[key], 0);
}

/**
 * the five scores of a response as text
 *
 * @param response  the response
 * @returns e.g. "Total 5 (Rules 1, Fouls 1, Fair-Mindedness 1, Attitude 1, Communication 1)"
 */
function _scoresText_(response: ResponseRecord): string {
  const scores = SCORE_KEYS.map((key) => `${RESPONSE_HEADERS[key]} ${response[key]}`).join(", ");
  return `Total ${_responseTotal_(response)} (${scores})`;
}

/**
 * where a form is, for Details
 *
 * @param r  the response
 * @returns e.g. "A1vA3" or "no game"
 */
function _gameText_(r: { game: string }): string {
  return r.game === "" ? "no game" : r.game;
}

/**
 * the Play and Scoring checks of every form
 * pure, no google calls
 *
 * Play checks look at counted forms only, so a duplicate or a form for no game is never held against the opponent
 * the comment checks skip international tournaments, comments are not available and the scorers are not UKU teams
 *
 * @param responses  every response, matched to games
 * @param deleted    every deleted form
 * @param events     events with Include Results ticked
 * @param settings   thresholds from the Issue Rules settings
 * @returns one finding per form per check
 */
function _gameFindings_(
  responses: ResponseRecord[],
  deleted: DeletedFormRecord[],
  events: EventRecord[],
  settings: IssueSettings,
): IssueFinding[] {
  const eventByUrl = new Map(events.map((e) => [e.spiritFormUrl, e]));
  const keywordChecks = COMMENT_KEYWORD_CHECKS.map((category) => ({
    category,
    pattern: _keywordPattern_(settings.commentKeywords[category]),
  }));
  const lowByReceiver = _lowScoresByReceiver_(responses, settings);
  const findings: IssueFinding[] = [];

  for (const r of responses) {
    const event = eventByUrl.get(r.spiritFormUrl);
    if (!event) continue;
    const comment = r.comment === "" ? null : { from: r.scorer, text: r.comment };
    const play = (category: IssueCategoryKey, text: string) =>
      findings.push({ category, event, team: r.receiver, game: r.game, opponent: r.scorer, text, comment });
    const scoring = (category: IssueCategoryKey, text: string) =>
      findings.push({ category, event, team: r.scorer, game: r.game, opponent: r.receiver, text, comment });

    if (r.match === MATCH_RESULTS.SELF_SCORE) {
      scoring("selfScore", `${r.scorer} gave itself ${_scoresText_(r)}`);
      continue;
    }
    if (r.match === MATCH_RESULTS.NO_SUCH_GAME) {
      scoring("noSuchGame", `${r.scorer} gave ${r.receiver} ${_scoresText_(r)}, but they have no game on the schedule`);
      continue;
    }
    if (r.match === MATCH_RESULTS.DUPLICATE) continue;

    const total = _responseTotal_(r);
    if (total <= settings.lowScoreAtOrBelow) {
      const low = lowByReceiver.get(`${r.spiritFormUrl}|${_teamKey_(r.receiver)}`) ?? [];
      play("lowScore", _lowScoreText_(r, low, settings));
    }
    const minimums = SCORE_KEYS.filter((key) => r[key] === settings.categoryMinimum);
    if (minimums.length > 0) {
      play(
        "categoryMinimum",
        `From ${r.scorer}: ${minimums.map((key) => `${RESPONSE_HEADERS[key]} ${r[key]}`).join(", ")}`,
      );
    }

    if (event.international) continue;

    if (r.comment === "") {
      if (total > settings.commentTotalAbove || total < settings.commentTotalBelow) {
        scoring("totalWithoutComment", `To ${r.receiver}: Total ${total}`);
      }
      const flagged = SCORE_KEYS.filter((key) => settings.commentCategoryScores.includes(r[key]));
      if (flagged.length > 0) {
        scoring(
          "categoryWithoutComment",
          `To ${r.receiver}: ${flagged.map((key) => `${RESPONSE_HEADERS[key]} ${r[key]}`).join(", ")}`,
        );
      }
    } else {
      for (const check of keywordChecks) {
        if (check.pattern && check.pattern.test(r.comment)) play(check.category, `From ${r.scorer}`);
      }
    }
  }

  findings.push(..._duplicateFindings_(responses, eventByUrl));

  for (const f of deleted) {
    const event = eventByUrl.get(f.spiritFormUrl);
    if (!event) continue;
    findings.push({
      category: "deletedRow",
      event,
      team: f.scorer,
      game: "",
      opponent: f.receiver,
      text: f.reason === "" ? "No reason given" : f.reason,
      comment: null,
    });
  }
  return findings;
}

/**
 * one Duplicate finding per scoring team → receiving team with a duplicate, listing every form for that pair
 * pure, never modifies its arguments
 *
 * @param responses   every response, matched to games
 * @param eventByUrl  Spirit Form URL → included event
 * @returns the findings
 */
function _duplicateFindings_(responses: ResponseRecord[], eventByUrl: Map<string, EventRecord>): IssueFinding[] {
  const pairs: Map<string, ResponseRecord[]> = new Map();
  for (const r of responses) {
    if (r.match !== MATCH_RESULTS.OK && r.match !== MATCH_RESULTS.DUPLICATE) continue;
    const key = `${r.spiritFormUrl}|${_teamKey_(r.scorer)}|${_teamKey_(r.receiver)}`;
    pairs.set(key, [...(pairs.get(key) ?? []), r]);
  }
  const findings: IssueFinding[] = [];
  for (const forms of pairs.values()) {
    const event = eventByUrl.get(forms[0].spiritFormUrl);
    if (!event || !forms.some((r) => r.match === MATCH_RESULTS.DUPLICATE)) continue;
    const lines = forms.map((r) =>
      `${r.timestamp ? _isoDateTime_(r.timestamp) : "no time"}: ${_scoresText_(r)}, `
      + `${r.counted ? `counted for ${_gameText_(r)}` : "duplicate"}${r.comment === "" ? "" : `\n"${r.comment}"`}`
    );
    findings.push({
      category: "duplicate",
      event,
      team: forms[0].scorer,
      game: "",
      opponent: forms[0].receiver,
      text: `${forms.length} forms to ${forms[0].receiver}:\n${lines.join("\n")}`,
      comment: null,
    });
  }
  return findings;
}

/**
 * every counted total at or below the low score threshold, by tournament and receiving team
 * pure, never modifies its arguments
 *
 * @param responses  every response
 * @param settings   thresholds
 * @returns "Spirit Form URL|receiver key" → its low scores in the order given
 */
function _lowScoresByReceiver_(responses: ResponseRecord[], settings: IssueSettings): Map<string, ResponseRecord[]> {
  const low: Map<string, ResponseRecord[]> = new Map();
  for (const r of responses) {
    if (!r.counted || _responseTotal_(r) > settings.lowScoreAtOrBelow) continue;
    const key = `${r.spiritFormUrl}|${_teamKey_(r.receiver)}`;
    low.set(key, [...(low.get(key) ?? []), r]);
  }
  return low;
}

/**
 * Details text of a low score: the score, the team's other low scores at the tournament, and whether policy is met
 *
 * @param r         the low score
 * @param allLow    every low score the team received at the tournament, including this one
 * @param settings  thresholds
 * @returns the text
 */
function _lowScoreText_(r: ResponseRecord, allLow: ResponseRecord[], settings: IssueSettings): string {
  const others = allLow.filter((x) => x !== r);
  const lines = [`From ${r.scorer}: ${_scoresText_(r)}`];
  if (others.length > 0) {
    lines.push(
      `Also ${settings.lowScoreAtOrBelow} or below at this tournament: `
        + others.map((x) => `${_responseTotal_(x)} from ${x.scorer} (${_gameText_(x)})`).join(", "),
    );
  }
  const met = allLow.length >= settings.lowScoreCount;
  lines.push(
    `${allLow.length} score(s) of ${settings.lowScoreAtOrBelow} or below at this tournament `
      + `(policy: ${settings.lowScoreCount} or more${met ? ", reached" : ""})`,
  );
  return lines.join("\n");
}

/**
 * one team's average at one tournament
 */
interface TeamAverage {
  /** the tournament */
  event: EventRecord;
  /** the team */
  team: string;
  /** counted totals it received */
  totals: number[];
  /** their average */
  average: number;
}

/**
 * every team's average at every tournament, from counted scores
 * pure, never modifies its arguments
 *
 * @param responses  every response
 * @param events     events with Include Results ticked
 * @returns one average per team per tournament
 */
function _teamAverages_(responses: ResponseRecord[], events: EventRecord[]): TeamAverage[] {
  const averages: TeamAverage[] = [];
  for (const event of events) {
    const received: Map<string, ResponseRecord[]> = new Map();
    for (const r of responses) {
      if (r.spiritFormUrl !== event.spiritFormUrl || !r.counted) continue;
      const key = _teamKey_(r.receiver);
      received.set(key, [...(received.get(key) ?? []), r]);
    }
    for (const forms of received.values()) {
      const totals = forms.map(_responseTotal_);
      const average = totals.reduce((a, b) => a + b, 0) / totals.length;
      averages.push({ event, team: forms[0].receiver, totals, average });
    }
  }
  return averages;
}

/**
 * Details text of an average issue: the average, then the club's earlier averages below either threshold
 * earlier means a tournament with an earlier start date
 *
 * @param average   the team's average at the tournament
 * @param all       every team's average at every tournament
 * @param clubOf    team key → club
 * @param settings  thresholds
 * @returns the text
 */
function _averageText_(
  average: TeamAverage,
  all: TeamAverage[],
  clubOf: Map<string, string>,
  settings: IssueSettings,
): string {
  const text = `Average ${average.average.toFixed(2)} from ${average.totals.length} scores: ${
    average.totals.join(", ")
  }`;
  const club = clubOf.get(_teamKey_(average.team)) ?? "";
  const start = average.event.startDate;
  if (club === "" || !start) return text;
  const earlier = all
    .filter((a) =>
      a.event.startDate !== null
      && a.event.startDate.getTime() < start.getTime()
      && clubOf.get(_teamKey_(a.team)) === club
      && a.average < settings.monitoringAverageBelow
    )
    .sort((a, b) => _compareDates_(a.event.startDate, b.event.startDate));
  if (earlier.length === 0) {
    return `${text}\n\nNo earlier averages below ${settings.monitoringAverageBelow} this season for ${club}`;
  }
  const lines = earlier.map((a) =>
    `- ${a.average < settings.immediateAverageBelow ? "Immediate" : "Monitoring"}: ${a.team} averaged `
    + `${a.average.toFixed(2)} at ${a.event.eventName} (${a.event.startDate ? _isoDate_(a.event.startDate) : ""})`
  );
  return `${text}\n\nEarlier this season for ${club}:\n${lines.join("\n")}`;
}

/**
 * the Tournament checks of every team at every tournament
 * pure, no google calls
 *
 * submission checks need the tournament's schedule, and skip international tournaments
 *   a team's games are the played games on the tournament's schedule,
 *   the teams are those in the tournament's responses, so divisions sharing a schedule stay apart
 * average checks use counted scores, a team gets at most one of Immediate and Monitoring,
 *   with the club's earlier averages below either threshold this season as context
 *
 * @param responses  every response, matched to games
 * @param games      every game, with who submitted
 * @param events     events with Include Results ticked, in date order
 * @param clubOf     team key → club, from the Teams tab
 * @param settings   thresholds from the Issue Rules settings
 * @returns one finding per team per check
 */
function _tournamentFindings_(
  responses: ResponseRecord[],
  games: GameRecord[],
  events: EventRecord[],
  clubOf: Map<string, string>,
  settings: IssueSettings,
): IssueFinding[] {
  const findings: IssueFinding[] = [];
  const averages = _teamAverages_(responses, events);

  for (const event of events) {
    const atEvent = responses.filter((r) => r.spiritFormUrl === event.spiritFormUrl);
    const teams: Map<string, string> = new Map();
    for (const r of atEvent) {
      if (!teams.has(_teamKey_(r.receiver))) teams.set(_teamKey_(r.receiver), r.receiver);
      if (!event.international && !teams.has(_teamKey_(r.scorer))) teams.set(_teamKey_(r.scorer), r.scorer);
    }
    const scheduled = atEvent.some((r) => r.match !== MATCH_RESULTS.NO_SCHEDULE);
    const eventGames = games.filter((g) => g.scheduleUrl === event.scheduleUrl && g.played);

    for (const [key, team] of teams) {
      const tournament = (category: IssueCategoryKey, text: string) =>
        findings.push({ category, event, team, game: "", opponent: "", text, comment: null });

      if (scheduled && !event.international) {
        const isTeam1 = (g: GameRecord) => _teamKey_(g.team1) === key;
        const played = eventGames.filter((g) => isTeam1(g) || _teamKey_(g.team2) === key);
        const missing = played.filter((g) => !(isTeam1(g) ? g.team1Submitted : g.team2Submitted));
        const submitted = played.length - missing.length;
        const percent = played.length === 0 ? 100 : (100 * submitted) / played.length;
        if (played.length > 0 && (submitted === 0 || percent < settings.submittedPercentBelow)) {
          const opponentOf = (g: GameRecord) => (isTeam1(g) ? g.team2 : g.team1);
          const text = `Submitted ${submitted} of ${played.length} games (${Math.round(percent)}%)\n\n`
            + `Missing for:\n${missing.map(opponentOf).join("\n")}\n\n`
            + `Games: ${missing.map((g) => `${g.game} v ${opponentOf(g)}`).join(", ")}`;
          tournament(submitted === 0 ? "noScores" : "underSubmitted", text);
        }
      }

      const average = averages.find((a) => a.event === event && _teamKey_(a.team) === key);
      if (!average) continue;
      if (average.average < settings.immediateAverageBelow) {
        tournament("immediateAverage", _averageText_(average, averages, clubOf, settings));
      } else if (average.average < settings.monitoringAverageBelow) {
        tournament("monitoringAverage", _averageText_(average, averages, clubOf, settings));
      }
    }
  }
  return findings;
}

/**
 * the issue id of a finding
 * readable, and the same every run, so a finding always joins the same issue
 *
 * @param finding  the finding
 * @returns e.g. "PLAY | 2026-10-31 NUXIR 2026-27 | A1vA3 | Durham 1", "SCORING | ... | v York 1 | Leeds 2"
 *   or "NO-SCORES | 2026-10-31 NUXIR 2026-27 | Durham 1"
 */
function _issueId_(finding: IssueFinding): string {
  const check = ISSUE_CATEGORIES[finding.category];
  const date = finding.event.startDate ? _isoDate_(finding.event.startDate) : "no date";
  const where = `${date} ${finding.event.eventName}`;
  if (check.level === ISSUE_LEVELS.TOURNAMENT) return `${check.code} | ${where} | ${finding.team}`;
  const game = finding.game === "" ? `v ${finding.opponent}` : finding.game;
  return `${check.level.toUpperCase()} | ${where} | ${game} | ${finding.team}`;
}

/**
 * Details sections for some findings of one issue: one per finding, then each comment quoted once
 * pure, never modifies its arguments
 *
 * @param findings  findings of the issue
 * @param settings  thresholds, for the labels
 * @returns the text
 */
function _findingsDetails_(findings: IssueFinding[], settings: IssueSettings): string {
  const sections = findings.map((f) =>
    `${ISSUE_CATEGORIES[f.category].code}: ${ISSUE_CATEGORIES[f.category].label(settings)}\n${f.text}`
  );
  const comments: Map<string, { from: string; text: string }> = new Map();
  for (const f of findings) if (f.comment) comments.set(`${f.comment.from}|${f.comment.text}`, f.comment);
  const quoted = [...comments.values()].map((c) => `Comment from ${c.from}:\n"${c.text}"`);
  return [...sections, ...quoted].join("\n\n");
}

/**
 * an issue found now, with its findings by check code
 */
interface FoundIssue {
  /** the issue as it would be added */
  issue: IssueRecord;
  /** check code → its findings */
  findings: Map<string, IssueFinding[]>;
}

/**
 * group findings into issues, one per Issue ID
 * pure, never modifies its arguments
 *
 * @param findings  every finding, from enabled checks
 * @param clubOf    team key → club, from the Teams tab
 * @param today     date to record as the Issue Date
 * @param settings  thresholds, for the labels
 * @returns issues in the order first found
 */
function _issuesFromFindings_(
  findings: IssueFinding[],
  clubOf: Map<string, string>,
  today: Date,
  settings: IssueSettings,
): FoundIssue[] {
  const byId: Map<string, IssueFinding[]> = new Map();
  for (const f of findings) {
    const id = _issueId_(f);
    byId.set(id, [...(byId.get(id) ?? []), f]);
  }
  return [...byId].map(([issueId, grouped]) => {
    const byCode: Map<string, IssueFinding[]> = new Map();
    for (const f of grouped) {
      const code = ISSUE_CATEGORIES[f.category].code;
      byCode.set(code, [...(byCode.get(code) ?? []), f]);
    }
    const first = grouped[0];
    return {
      issue: {
        issueId,
        dateCreated: today,
        level: ISSUE_CATEGORIES[first.category].level,
        checks: [...byCode.keys()].join(", "),
        tournament: first.event.eventName,
        tournamentDate: first.event.startDate,
        game: first.game,
        opponent: first.opponent,
        team: first.team,
        club: clubOf.get(_teamKey_(first.team)) ?? "",
        details: `${first.event.international ? "INTERNATIONAL\n\n" : ""}${_findingsDetails_(grouped, settings)}`,
        status: ISSUE_STATUSES[0],
        email: "",
        committeeMember: "",
        notes: "",
      },
      findings: byCode,
    };
  });
}

/**
 * the codes in a Checks cell
 *
 * @param checks  e.g. "LOW-SCORE, DANGEROUS-PLAY"
 * @returns the codes
 */
function _checkCodes_(checks: unknown): string[] {
  return String(checks ?? "").split(",").map((code) => code.trim()).filter((code) => code !== "");
}

/**
 * the Checks, Details and Status of existing issues after adding newly found checks
 * pure, never modifies its arguments
 * an issue with a new check gets it added to Checks and its section to Details, and goes back to NEW if CLOSED
 *
 * @param existing  existing issues by row: Issue ID, Checks, Details and Status
 * @param found     Issue ID → findings by check code, for every issue found now
 * @param settings  thresholds, for the labels
 * @returns the three columns for every existing row, and how many issues changed
 */
function _mergeIssues_(
  existing: { issueId: string; checks: string; details: string; status: string }[],
  found: Map<string, Map<string, IssueFinding[]>>,
  settings: IssueSettings,
): { rows: { checks: string; details: string; status: string }[]; updated: number } {
  let updated = 0;
  const rows = existing.map((row) => {
    const findings = found.get(row.issueId);
    if (!findings) return row;
    const have = new Set(_checkCodes_(row.checks));
    const added = [...findings].filter(([code]) => !have.has(code));
    if (added.length === 0) return row;
    updated++;
    return {
      checks: [..._checkCodes_(row.checks), ...added.map(([code]) => code)].join(", "),
      details: `${row.details}\n\n${_findingsDetails_(added.flatMap(([, f]) => f), settings)}`,
      status: row.status.trim().toUpperCase() === "CLOSED" ? ISSUE_STATUSES[0] : row.status,
    };
  });
  return { rows, updated };
}

/**
 * @returns the Issues tab, created on first use
 */
function _getIssuesSheet_(): GoogleAppsScript.Spreadsheet.Sheet {
  const headers: readonly string[] = Object.values(ISSUE_HEADERS);
  return _getOrCreateSheet_(SpreadsheetApp.getActiveSpreadsheet(), ISSUES_SHEET, headers, ISSUES_INFO);
}

/**
 * add new issues at the bottom, and add newly found checks to existing ones
 * existing rows only ever have Checks, Details and Status written, see _mergeIssues_
 *
 * @param found     every issue found now
 * @param settings  thresholds, for the labels
 * @returns how many issues were added and how many were updated
 */
function _writeIssues_(found: FoundIssue[], settings: IssueSettings): { added: number; updated: number } {
  const sheet = _getIssuesSheet_();
  const lastRow = sheet.getLastRow();
  const column = (key: keyof typeof ISSUE_HEADERS) => ISSUE_KEYS.indexOf(key) + 1;
  const rowCount = Math.max(lastRow - HEADER_ROW, 0);
  const read = (key: keyof typeof ISSUE_HEADERS) =>
    rowCount === 0 ? [] : sheet.getRange(DATA_ROW, column(key), rowCount, 1).getValues().map((r) => String(r[0]));
  const ids = read("issueId");
  const checks = read("checks");
  const details = read("details");
  const statuses = read("status");

  const merged = _mergeIssues_(
    ids.map((issueId, i) => ({ issueId, checks: checks[i], details: details[i], status: statuses[i] })),
    new Map(found.map((f) => [f.issue.issueId, f.findings])),
    settings,
  );
  if (merged.updated > 0) {
    for (const key of ["checks", "details", "status"] as const) {
      sheet.getRange(DATA_ROW, column(key), rowCount, 1).setValues(merged.rows.map((r) => [r[key]]));
    }
  }

  const existingIds = new Set(ids);
  const added = found.map((f) => f.issue).filter((issue) => !existingIds.has(issue.issueId));
  if (added.length > 0) {
    const firstNew = Math.max(lastRow, HEADER_ROW) + 1;
    _fitSheet_(sheet, firstNew + added.length - 1, ISSUE_KEYS.length);
    sheet.getRange(firstNew, 1, added.length, ISSUE_KEYS.length).setValues(
      added.map((issue) => ISSUE_KEYS.map((key) => issue[key] ?? "")),
    );
    const statusRule = SpreadsheetApp.newDataValidation().requireValueInList([...ISSUE_STATUSES], true).build();
    sheet.getRange(firstNew, column("status"), added.length, 1).setDataValidation(statusRule);
    for (const key of ["dateCreated", "tournamentDate"] as const) {
      sheet.getRange(firstNew, column(key), added.length, 1).setNumberFormat("yyyy-mm-dd");
    }
    _applyFilter_(sheet, firstNew + added.length - 1 - HEADER_ROW, ISSUE_KEYS.length);
    _styleRows_(sheet, firstNew, added.length, ISSUE_KEYS.length, ISSUE_EDITABLE);
  }
  _styleRows_(sheet, HEADER_ROW, 1, ISSUE_KEYS.length, ISSUE_EDITABLE);
  return { added: added.length, updated: merged.updated };
}

/**
 * run every check on what is already in the spreadsheet, then add new issues and new checks to existing issues
 * no files are read, use after Refresh Results, or after changing Include Results, International,
 * a club override or an issue rule
 * reads the clubs from the Teams tab and the thresholds from the Issue Rules settings, disabled checks add nothing
 *
 * @returns a one-line summary
 */
function _refreshIssues_(): string {
  const { events, responses, games, deleted, settings, clubOf } = _timed_("read for issues", () => {
    const events = _sortEvents_(_readEvents_(_getEventsSheet_())).filter((e) => e.includeResults);
    return {
      events,
      responses: _sortResponses_(_readResponses_(_getResponsesSheet_()), events),
      games: _readGames_(_getGamesSheet_()),
      deleted: _readDeletedForms_(_getDeletedFormsSheet_()),
      settings: _readIssueSettings_(),
      clubOf: _readTeamClubs_(),
    };
  });
  if (responses.length === 0) return "No results found: run Refresh Results first";

  const found = _timed_("check issues", () => {
    const findings = [
      ..._gameFindings_(responses, deleted, events, settings),
      ..._tournamentFindings_(responses, games, events, clubOf, settings),
    ].filter((f) => settings.enabled[f.category]);
    return _issuesFromFindings_(findings, clubOf, new Date(), settings);
  });
  const { added, updated } = _timed_("write issues", () => _writeIssues_(found, settings));
  const emails = _timed_("email links", () => {
    try {
      return _updateEmailLinks_();
    } catch (e) {
      return `email links not updated: ${e instanceof Error ? e.message : String(e)}`;
    }
  });
  return `${added} new issue(s), ${updated} updated; ${emails}`;
}
