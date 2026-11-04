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
  + "Club looks the team up on Teams, so it follows Club Overrides. "
  + "Email Address looks the club up on Club Contacts. Email Subject and Email Body are only filled while the issue "
  + "is NEW, ready to copy into an email. Changing Status clears them straight away, setting it back to NEW "
  + "fills them again\n\n"
  + "A refresh that finds another check for an existing issue adds it to Checks and Details, "
  + "and a CLOSED issue goes back to NEW. Nothing else on existing rows is changed\n\n"
  + "Refresh Results and Issues puts the tab back in order: club, tournament, team, level, checks, opponent. "
  + "Filters are kept, so filter freely";

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
    club: "Club",
    team: "Team",
    tournament: "Tournament",
    tournamentDate: "Tournament Date",
    game: "Game",
    opponent: "Opponent",
    score: "Score",
    spirit: "Spirit",
    comment: "Comment",
    checks: "Checks",
    details: "Details",
    status: "Status",
    committeeMember: "Committee Member",
    notes: "Notes",
    emailAddress: "Email Address",
    emailSubject: "Email Subject",
    emailBody: "Email Body",
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
 * the order of the Issues tab, put back after every refresh:
 * a club's issues together, then by tournament (date, then name), team, level, checks and opponent
 */
const ISSUE_SORT: readonly (keyof typeof ISSUE_HEADERS)[] = Object.freeze(
  ["club", "tournamentDate", "tournament", "team", "level", "checks", "opponent"] as const,
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
 */
const ISSUE_CATEGORIES = Object.freeze({
  lowScore: { code: "LOW-SCORE", level: ISSUE_LEVELS.PLAY },
  dangerousPlay: { code: "DANGEROUS-PLAY", level: ISSUE_LEVELS.PLAY },
  cheating: { code: "CHEATING", level: ISSUE_LEVELS.PLAY },
  harassment: { code: "HARASSMENT", level: ISSUE_LEVELS.PLAY },
  categoryMinimum: { code: "MIN-CATEGORY", level: ISSUE_LEVELS.PLAY },
  totalWithoutComment: { code: "TOTAL-NO-COMMENT", level: ISSUE_LEVELS.SCORING },
  categoryWithoutComment: { code: "CATEGORY-NO-COMMENT", level: ISSUE_LEVELS.SCORING },
  deletedRow: { code: "DELETED-ROW", level: ISSUE_LEVELS.SCORING },
  redactedComment: { code: "REDACTED-COMMENT", level: ISSUE_LEVELS.SCORING },
  selfScore: { code: "SELF-SCORE", level: ISSUE_LEVELS.SCORING },
  noSuchGame: { code: "NO-SUCH-GAME", level: ISSUE_LEVELS.SCORING },
  duplicate: { code: "DUPLICATE", level: ISSUE_LEVELS.SCORING },
  noScores: { code: "NO-SCORES", level: ISSUE_LEVELS.TOURNAMENT },
  underSubmitted: { code: "UNDER-SUBMITTED", level: ISSUE_LEVELS.TOURNAMENT },
  immediateAverage: { code: "IMMEDIATE-AVERAGE", level: ISSUE_LEVELS.TOURNAMENT },
  monitoringAverage: { code: "MONITORING-AVERAGE", level: ISSUE_LEVELS.TOURNAMENT },
});

/**
 * how each comment keyword check is named in Details, e.g. Comment includes dangerous play keyword: "danger"
 */
const KEYWORD_CHECK_NAMES: Readonly<Record<typeof COMMENT_KEYWORD_CHECKS[number], string>> = Object.freeze({
  dangerousPlay: "dangerous play",
  cheating: "cheating",
  harassment: "harassment",
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
  /** the game's result from the team's side, e.g. 15-12W, "" without a game or scores */
  score: string;
  /** the spirit score behind the issue, e.g. "10 (2, 2, 2, 2, 2)": received for Play, given for Scoring */
  spirit: string;
  /** the comment behind the issue, in quotes, from the same form as Spirit */
  comment: string;
  /** team the issue is about */
  team: string;
  /** club of the team, written as a formula looking it up on Teams, see _issueClubFormula_ */
  club: string;
  /** one section per check */
  details: string;
  /** one of ISSUE_STATUSES */
  status: string;
  /** formula looking up the club's email on Club Contacts, see _emailAddressFormula_ */
  emailAddress: string;
  /** filled from the email template while the issue is NEW, see _updateEmailDrafts_ */
  emailSubject: string;
  /** filled from the email template while the issue is NEW, see _updateEmailDrafts_ */
  emailBody: string;
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
  /** what was found, for Details, "" when Checks says it all */
  text: string;
  /** the form's comment for the Comment column, in quotes, "" for none */
  comment: string;
  /** the form's spirit score, see _spiritText_, "" when there is not one form */
  spirit: string;
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
 * a form's spirit score for the Spirit column: the total, then each category in form order
 * pure, no google calls
 *
 * @param scores  the five category scores, null where missing
 * @returns e.g. "10 (2, 2, 2, 2, 2)", "" if any score is missing
 */
function _spiritText_(scores: Record<typeof SCORE_KEYS[number], number | null>): string {
  const values = SCORE_KEYS.map((key) => scores[key]);
  if (values.some((v) => v === null)) return "";
  return `${values.reduce((sum: number, v) => sum + (v ?? 0), 0)} (${values.join(", ")})`;
}

/**
 * a comment in quotes, for the Comment column and Details
 *
 * @param comment  the comment
 * @returns the comment in double quotes, "" for no comment
 */
function _quoted_(comment: string): string {
  return comment === "" ? "" : `"${comment}"`;
}

/**
 * names joined as readable text, e.g. "Rules, Fouls and Attitude"
 *
 * @param names  the names
 * @returns the text
 */
function _andList_(names: string[]): string {
  return names.length <= 1 ? names.join("") : `${names.slice(0, -1).join(", ")} and ${names[names.length - 1]}`;
}

/**
 * a category score as words, e.g. "a zero", "a 4"
 *
 * @param score  the score
 * @returns the text
 */
function _scoreWords_(score: number): string {
  return score === 0 ? "a zero" : `a ${score}`;
}

/**
 * which of the words a comment includes, as whole words ignoring case
 *
 * @param comment  the comment
 * @param words    the words of one keyword check
 * @returns the words found, in list order
 */
function _keywordsFound_(comment: string, words: string[]): string[] {
  return words.filter((word) => _keywordPattern_([word])?.test(comment));
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
 * @param games      every game, to put deleted forms in a game
 * @param events     events with Include Results ticked
 * @param settings   thresholds from the Issue Rules settings
 * @returns one finding per form per check
 */
function _gameFindings_(
  responses: ResponseRecord[],
  deleted: DeletedFormRecord[],
  games: GameRecord[],
  events: EventRecord[],
  settings: IssueSettings,
): IssueFinding[] {
  const eventByUrl = new Map(events.map((e) => [e.spiritFormUrl, e]));
  const lowByReceiver = _lowScoresByReceiver_(responses, settings);
  const findings: IssueFinding[] = [];

  for (const r of responses) {
    const event = eventByUrl.get(r.spiritFormUrl);
    if (!event) continue;
    const comment = _quoted_(r.comment);
    const spirit = _spiritText_(r);
    const play = (category: IssueCategoryKey, text: string) =>
      findings.push({ category, event, team: r.receiver, game: r.game, opponent: r.scorer, text, comment, spirit });
    const scoring = (category: IssueCategoryKey, text: string) =>
      findings.push({ category, event, team: r.scorer, game: r.game, opponent: r.receiver, text, comment, spirit });

    if (r.match === MATCH_RESULTS.SELF_SCORE) {
      scoring("selfScore", "");
      continue;
    }
    if (r.match === MATCH_RESULTS.NO_SUCH_GAME) {
      scoring("noSuchGame", `No game against ${r.receiver} on the schedule`);
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
        `Received ${_scoreWords_(settings.categoryMinimum)} in ${
          _andList_(minimums.map((key) => RESPONSE_HEADERS[key]))
        }`,
      );
    }

    if (event.international) continue;

    if (r.comment === "") {
      const totalText = _totalWithoutCommentText_(r, settings);
      if (totalText !== "") scoring("totalWithoutComment", totalText);
      const categoryText = _categoryWithoutCommentText_(r, settings);
      if (categoryText !== "") scoring("categoryWithoutComment", categoryText);
    } else {
      for (const category of COMMENT_KEYWORD_CHECKS) {
        const found = _keywordsFound_(r.comment, settings.commentKeywords[category]);
        if (found.length === 0) continue;
        const words = found.map((word) => `"${word}"`).join(", ");
        play(
          category,
          `Comment includes ${KEYWORD_CHECK_NAMES[category]} keyword${found.length > 1 ? "s" : ""}: ${words}`,
        );
      }
    }
  }

  findings.push(..._duplicateFindings_(responses, eventByUrl));
  findings.push(..._deletedFindings_(deleted, responses, games, eventByUrl, settings));
  return findings;
}

/**
 * Details text when a form's total needs a comment, from the TOTAL-NO-COMMENT thresholds
 * the form's comment is not looked at, callers only ask about forms without one
 *
 * @param scores    the five category scores, null where missing
 * @param settings  thresholds
 * @returns e.g. "Gave a total of 16 without a comment, ...", "" if the total does not need a comment or is incomplete
 */
function _totalWithoutCommentText_(
  scores: Record<typeof SCORE_KEYS[number], number | null>,
  settings: IssueSettings,
): string {
  const values = SCORE_KEYS.map((key) => scores[key]);
  if (values.some((v) => v === null)) return "";
  const total = values.reduce((sum: number, v) => sum + (v ?? 0), 0);
  if (total <= settings.commentTotalAbove && total >= settings.commentTotalBelow) return "";
  return `Gave a total of ${total} without a comment, `
    + `comment required above ${settings.commentTotalAbove} or below ${settings.commentTotalBelow}`;
}

/**
 * Details text when a form's category scores need a comment, from the CATEGORY-NO-COMMENT scores
 * the form's comment is not looked at, callers only ask about forms without one
 *
 * @param scores    the five category scores, null where missing
 * @param settings  thresholds
 * @returns e.g. "Gave a 4 in Rules and Fouls without a comment", "" if no category needs a comment
 */
function _categoryWithoutCommentText_(
  scores: Record<typeof SCORE_KEYS[number], number | null>,
  settings: IssueSettings,
): string {
  const flagged = SCORE_KEYS.filter((key) => {
    const score = scores[key];
    return score !== null && settings.commentCategoryScores.includes(score);
  });
  if (flagged.length === 0) return "";
  const byScore = [...new Set(flagged.map((key) => scores[key] as number))].map((score) =>
    `${_scoreWords_(score)} in ${
      _andList_(flagged.filter((key) => scores[key] === score).map((key) => RESPONSE_HEADERS[key]))
    }`
  );
  return `Gave ${_andList_(byScore)} without a comment`;
}

/**
 * whether a Deleted forms reason says the form was moved for needing a comment
 * the Schedule Manager's delete step writes "Total<6 & Comment required", "Total>14 & Comment required"
 * or "Comment required", the check codes TOTAL-NO-COMMENT and CATEGORY-NO-COMMENT are recognised too
 *
 * @param reason  the "Reason for deletion" cell
 * @returns true for any of those, ignoring case and surrounding spaces
 */
function _isCommentRequiredReason_(reason: string): boolean {
  const text = reason.trim().toLowerCase();
  return text.endsWith("comment required")
    || [ISSUE_CATEGORIES.totalWithoutComment.code, ISSUE_CATEGORIES.categoryWithoutComment.code]
      .some((code) => text === code.toLowerCase());
}

/**
 * one finding per deleted form, or one per comment check it breaks
 * pure, never modifies its arguments
 *
 * a deleted form whose teams and submission time match a form still on Results, with a different comment,
 * is the original of a redacted comment: REDACTED-COMMENT, in the game of the form on Results, Comment shows both
 * a deleted form whose reason says it needed a comment raises TOTAL-NO-COMMENT and CATEGORY-NO-COMMENT
 * as if it were still on Results, falling back to DELETED-ROW if its scores break neither
 * any other deleted form was taken out: DELETED-ROW, Comment shows its comment
 * forms not on Results are put in a game when their two teams played only once there
 *
 * @param deleted     every deleted form
 * @param responses   every response, to find edited comments
 * @param games       every game, to find the game of a deleted form
 * @param eventByUrl  Spirit Form URL → included event
 * @param settings    thresholds from the Issue Rules settings
 * @returns the findings
 */
function _deletedFindings_(
  deleted: DeletedFormRecord[],
  responses: ResponseRecord[],
  games: GameRecord[],
  eventByUrl: Map<string, EventRecord>,
  settings: IssueSettings,
): IssueFinding[] {
  const formKey = (f: { spiritFormUrl: string; scorer: string; receiver: string; timestamp: Date | null }) =>
    `${f.spiritFormUrl}|${_teamKey_(f.scorer)}|${_teamKey_(f.receiver)}|${f.timestamp?.getTime() ?? ""}`;
  const kept = new Map(responses.filter((r) => r.timestamp !== null).map((r) => [formKey(r), r]));
  const onlyGame = (event: EventRecord, a: string, b: string) => {
    const between = games.filter((g) =>
      g.scheduleUrl === event.scheduleUrl
      && [_teamKey_(g.team1), _teamKey_(g.team2)].sort().join("|") === [_teamKey_(a), _teamKey_(b)].sort().join("|")
    );
    return between.length === 1 ? between[0].game : "";
  };

  const findings: IssueFinding[] = [];
  for (const f of deleted) {
    const event = eventByUrl.get(f.spiritFormUrl);
    if (!event) continue;
    const reason = f.reason === "" ? "no reason given" : f.reason;
    const spirit = _spiritText_(f);
    const edited = f.timestamp === null ? undefined : kept.get(formKey(f));
    const base = { event, team: f.scorer, opponent: f.receiver, spirit };

    if (edited !== undefined && edited.comment !== f.comment) {
      findings.push({
        ...base,
        category: "redactedComment",
        game: edited.game,
        text: `Comment redacted: ${reason}`,
        comment: `Original comment: ${_quoted_(f.comment) || "\"\""}\n\nRedacted comment: ${
          _quoted_(edited.comment) || "\"\""
        }`,
      });
      continue;
    }

    const game = onlyGame(event, f.scorer, f.receiver);
    const comment = _quoted_(f.comment);
    if (_isCommentRequiredReason_(f.reason) && f.comment === "" && !event.international) {
      const totalText = _totalWithoutCommentText_(f, settings);
      const categoryText = _categoryWithoutCommentText_(f, settings);
      if (totalText !== "") findings.push({ ...base, category: "totalWithoutComment", game, text: totalText, comment });
      if (categoryText !== "") {
        findings.push({ ...base, category: "categoryWithoutComment", game, text: categoryText, comment });
      }
      if (totalText !== "" || categoryText !== "") continue;
    }
    findings.push({ ...base, category: "deletedRow", game, text: `Deleted row: ${reason}`, comment });
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
    const lines = forms.map((r) => [_spiritText_(r), _quoted_(r.comment)].filter((x) => x !== "").join(" "));
    findings.push({
      category: "duplicate",
      event,
      team: forms[0].scorer,
      game: "",
      opponent: forms[0].receiver,
      text: "More responses submitted than games played.\n\n"
        + `Games played: ${forms.filter((r) => r.counted).length}.\n\n`
        + `Responses:\n${lines.join("\n")}`,
      comment: "",
      spirit: "",
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
 * Details text of a low score, with the team's other low scores at the tournament when it has any,
 * as two or more meets the policy
 *
 * @param r         the low score
 * @param allLow    every low score the team received at the tournament, including this one
 * @param settings  thresholds
 * @returns the text
 */
function _lowScoreText_(r: ResponseRecord, allLow: ResponseRecord[], settings: IssueSettings): string {
  const text = `Received a total of ${settings.lowScoreAtOrBelow} or below`;
  const others = allLow.filter((x) => x !== r);
  if (others.length === 0) return text;
  const list = others.map((x) => `${_responseTotal_(x)} from ${x.scorer}${x.game === "" ? "" : ` (${x.game})`}`);
  const met = allLow.length >= settings.lowScoreCount ? `, ${settings.lowScoreCount} or more is policy` : "";
  return `${text}\n\nAlso received ${list.join(", ")}: ${allLow.length} at this tournament${met}`;
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
  const below = average.average < settings.immediateAverageBelow
    ? settings.immediateAverageBelow
    : settings.monitoringAverageBelow;
  const text = `Average of ${average.average.toFixed(2)} from ${average.totals.length} scores, below ${below}`
    + `\n\nScores: ${average.totals.join(", ")}`;
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
        findings.push({ category, event, team, game: "", opponent: "", text, comment: "", spirit: "" });

      if (scheduled && !event.international) {
        const isTeam1 = (g: GameRecord) => _teamKey_(g.team1) === key;
        const played = eventGames.filter((g) => isTeam1(g) || _teamKey_(g.team2) === key);
        const missing = played.filter((g) => !(isTeam1(g) ? g.team1Submitted : g.team2Submitted));
        const submitted = played.length - missing.length;
        const percent = played.length === 0 ? 100 : (100 * submitted) / played.length;
        if (played.length > 0 && (submitted === 0 || percent < settings.submittedPercentBelow)) {
          const opponentOf = (g: GameRecord) => (isTeam1(g) ? g.team2 : g.team1);
          const done = played.filter((g) => !missing.includes(g)).map(opponentOf);
          const missingList = `Missing: ${missing.map(opponentOf).join(", ")}`;
          if (submitted === 0) {
            tournament("noScores", `0/${played.length} spirit scores submitted\n\n${missingList}`);
          } else {
            tournament(
              "underSubmitted",
              `${submitted}/${played.length} spirit scores submitted, ${Math.round(percent)}%. `
                + `Less than ${settings.submittedPercentBelow}%\n\nSubmitted: ${done.join(", ")}\n\n${missingList}`,
            );
          }
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
 * Details for some findings of one issue: each finding's text, a blank line between them
 * pure, never modifies its arguments
 *
 * @param findings  findings of the issue
 * @returns the text
 */
function _findingsDetails_(findings: IssueFinding[]): string {
  return [...new Set(findings.map((f) => f.text).filter((text) => text !== ""))].join("\n\n");
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
 * a game's result from one team's side
 * pure, no google calls
 *
 * @param game  the game, undefined when there is none
 * @param team  the team the result is for
 * @returns e.g. "15-12W", "12-15L" or "10-10D", "" without a game or both scores
 */
function _scoreText_(game: GameRecord | undefined, team: string): string {
  if (!game || game.score1 === null || game.score2 === null) return "";
  const isTeam1 = _teamKey_(game.team1) === _teamKey_(team);
  const own = isTeam1 ? game.score1 : game.score2;
  const other = isTeam1 ? game.score2 : game.score1;
  return `${own}-${other}${own > other ? "W" : own < other ? "L" : "D"}`;
}

/**
 * group findings into issues, one per Issue ID
 * pure, never modifies its arguments
 *
 * @param findings  every finding, from enabled checks
 * @param games     every game, for the Score column
 * @param clubOf    team key → club, from the Teams tab
 * @param today     date to record as the Issue Date
 * @returns issues in the order first found
 */
function _issuesFromFindings_(
  findings: IssueFinding[],
  games: GameRecord[],
  clubOf: Map<string, string>,
  today: Date,
): FoundIssue[] {
  const gameOf = new Map(games.map((g) => [`${g.scheduleUrl}|${g.game}`, g]));
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
        score: first.game === "" ? "" : _scoreText_(gameOf.get(`${first.event.scheduleUrl}|${first.game}`), first.team),
        spirit: grouped.find((f) => f.spirit !== "")?.spirit ?? "",
        comment: grouped.find((f) => f.comment !== "")?.comment ?? "",
        team: first.team,
        club: clubOf.get(_teamKey_(first.team)) ?? "",
        details: `${first.event.international ? "INTERNATIONAL\n\n" : ""}${_findingsDetails_(grouped)}`,
        status: ISSUE_STATUSES[0],
        emailAddress: "",
        emailSubject: "",
        emailBody: "",
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
 * @returns the three columns for every existing row, and how many issues changed
 */
function _mergeIssues_(
  existing: { issueId: string; checks: string; details: string; status: string }[],
  found: Map<string, Map<string, IssueFinding[]>>,
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
      details: [row.details, _findingsDetails_(added.flatMap(([, f]) => f))].filter((t) => t !== "").join("\n\n"),
      status: row.status.trim().toUpperCase() === "CLOSED" ? ISSUE_STATUSES[0] : row.status,
    };
  });
  return { rows, updated };
}

/**
 * formula for the Club cell of one Issues row: the team's club on the Teams tab
 * a formula, so a Club Override or Name Rule change shows on every issue straight away
 *
 * @param row  1-based sheet row the formula is for
 * @returns the formula
 */
function _issueClubFormula_(row: number): string {
  const team = `$${_columnLetter_(ISSUE_KEYS.indexOf("team") + 1)}${row}`;
  const teams = (key: keyof typeof TEAM_HEADERS) => _columnBelowHeader_(TEAMS_SHEET, TEAM_KEYS.indexOf(key) + 1);
  return `=IFERROR(XLOOKUP(${team}, ${teams("team")}, ${teams("club")}), "")`;
}

/**
 * @returns the Issues tab, created on first use
 */
function _getIssuesSheet_(): GoogleAppsScript.Spreadsheet.Sheet {
  const headers: readonly string[] = Object.values(ISSUE_HEADERS);
  return _getOrCreateSheet_(SpreadsheetApp.getActiveSpreadsheet(), ISSUES_SHEET, headers, ISSUES_INFO);
}

/**
 * add new issues, add newly found checks to existing ones, then put the tab back in ISSUE_SORT order
 * existing rows only ever have Checks, Details and Status written, see _mergeIssues_
 *
 * @param found  every issue found now
 * @returns how many issues were added and how many were updated
 */
function _writeIssues_(found: FoundIssue[]): { added: number; updated: number } {
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
      added.map((issue, i) =>
        ISSUE_KEYS.map((key) => {
          if (key === "club") return _issueClubFormula_(firstNew + i);
          if (key === "emailAddress") return _emailAddressFormula_(firstNew + i);
          return issue[key] ?? "";
        })
      ),
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
  const total = rowCount + added.length;
  if (total > 1) {
    sheet
      .getRange(DATA_ROW, 1, total, ISSUE_KEYS.length)
      .sort(ISSUE_SORT.map((key) => ({ column: column(key), ascending: true })));
  }
  return { added: added.length, updated: merged.updated };
}

/**
 * run every check on what is already in the spreadsheet, then add new issues and new checks to existing issues
 * no files are read, run by Refresh Results and Issues straight after importing
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
  if (responses.length === 0) return "No results to check: tick Include Results on the Tournaments tab";

  const found = _timed_("check issues", () => {
    const findings = [
      ..._gameFindings_(responses, deleted, games, events, settings),
      ..._tournamentFindings_(responses, games, events, clubOf, settings),
    ].filter((f) => settings.enabled[f.category]);
    return _issuesFromFindings_(findings, games, clubOf, new Date());
  });
  const { added, updated } = _timed_("write issues", () => _writeIssues_(found));
  const emails = _timed_("email drafts", () => {
    try {
      return _updateEmailDrafts_();
    } catch (e) {
      return `email drafts not updated: ${e instanceof Error ? e.message : String(e)}`;
    }
  });
  return `${added} new issue(s), ${updated} updated; ${emails}`;
}
