/**
 * name of the Issues tab
 */
const ISSUES_SHEET = "Issues";

/**
 * text for the info row above the Issues table
 * what the tab is, what to edit, how it refreshes
 */
const ISSUES_INFO = "Spirit issues found in the results, one row per issue per team per tournament\n\n"
  + "User editable columns:\n"
  + "- Status: NEW, IN PROGRESS or CLOSED\n"
  + "- Committee Member: who is handling it\n"
  + "- Notes\n\n"
  + "New issues are added by Refresh Issues, existing rows are never changed, so sort and filter freely";

/**
 * issues table columns, in order
 * keys are the names used in code
 * values are the header text
 */
const ISSUE_HEADERS = Object.freeze(
  {
    issueId: "Issue ID",
    dateCreated: "Issue Date",
    club: "Club",
    category: "Issue Category",
    team: "Team",
    tournaments: "Tournament(s)",
    tournamentDate: "Tournament Date",
    details: "Details",
    received: "Received",
    given: "Given",
    tournamentAverage: "Tournament Average",
    status: "Status",
    committeeMember: "Committee Member",
    notes: "Notes",
  } as const,
);

/**
 * keys of the Issues tab in column order
 */
const ISSUE_KEYS = Object.keys(ISSUE_HEADERS) as (keyof typeof ISSUE_HEADERS)[];

/**
 * values of the Status column, first is given to new issues
 */
const ISSUE_STATUSES = Object.freeze(["NEW", "IN PROGRESS", "CLOSED"]);

/**
 * every issue check
 * code goes into the Issue ID and the Rule column of the Issue Rules settings
 * label goes into the Issue Category column, built from the current settings
 */
const ISSUE_CATEGORIES = Object.freeze({
  totalWithoutComment: {
    code: "TOTAL-NO-COMMENT",
    label: (s: IssueSettings) => `Total above ${s.commentTotalAbove} or below ${s.commentTotalBelow} without comment`,
  },
  categoryWithoutComment: {
    code: "CATEGORY-NO-COMMENT",
    label: (s: IssueSettings) => `Category scored ${s.commentCategoryScores.join(" or ")} without comment`,
  },
  dangerousPlay: {
    code: "DANGEROUS-PLAY",
    label: (s: IssueSettings) => "Dangerous play mentioned",
  },
  cheating: {
    code: "CHEATING",
    label: (s: IssueSettings) => "Cheating mentioned",
  },
  harassment: {
    code: "HARASSMENT",
    label: (s: IssueSettings) => "Harassment mentioned",
  },
  notSubmitted: {
    code: "NOT-SUBMITTED",
    label: (s: IssueSettings) => "Spirit scores not submitted",
  },
  twoLowScores: {
    code: "LOW-SCORES",
    label: (s: IssueSettings) => `${s.lowScoreCount} or more scores of ${s.lowScoreAtOrBelow} or below at a tournament`,
  },
  lowAverage: {
    code: "LOW-AVERAGE",
    label: (s: IssueSettings) => `Average score below ${s.lowAverageBelow} at a tournament`,
  },
  categoryMinimum: {
    code: "MIN-CATEGORY",
    label: (s: IssueSettings) => `Received ${s.categoryMinimum} in a category`,
  },
  singleLowScore: {
    code: "SINGLE-LOW-SCORE",
    label: (s: IssueSettings) => `Received a score below ${s.singleLowScoreBelow}`,
  },
  monitoring: {
    code: "MONITORING",
    label: (s: IssueSettings) =>
      `${s.monitoringBreaches} team averages below ${s.monitoringAverageBelow} in the season`,
  },
  monitoringBreach: {
    code: "MONITORING-BREACH",
    label: (s: IssueSettings) => `Averaged below ${s.monitoringAverageBelow} again while on monitoring`,
  },
});

/**
 * checks raised by words in a comment, each with its own word list in the Issue Rules settings
 * the issue is for the receiving team and its Details quote the comment
 * a comment matching more than one list raises one issue per check
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
 * key of one issue check, e.g. "lowAverage"
 */
type IssueCategoryKey = keyof typeof ISSUE_CATEGORIES;

/**
 * one row of the Issues tab
 */
interface IssueRecord {
  /** check, tournament and team, unique per issue */
  issueId: string;
  /** when the issue was first added */
  dateCreated: Date;
  /** club the issue is about */
  club: string;
  /** label of the check */
  category: string;
  /** team the issue is about */
  team: string;
  /** tournament name(s) */
  tournaments: string;
  /** tournament date, null if unknown, comma separated text for several */
  tournamentDate: Date | string | null;
  /** one line per response that triggered it */
  details: string;
  /** total of each flagged response, comma separated */
  received: string;
  /** total of the reply to each flagged response, comma separated */
  given: string;
  /** average received at the tournament by each team that received a flagged score */
  tournamentAverage: string;
  /** one of ISSUE_STATUSES */
  status: string;
  /** who is handling it */
  committeeMember: string;
  /** free text */
  notes: string;
}

/**
 * one response that triggered a check, before grouping into issues
 */
interface IssueHit {
  /** which check */
  category: IssueCategoryKey;
  /** the event the response is from */
  event: EventRecord;
  /** team the issue is about */
  team: string;
  /** the other team in the response */
  other: string;
  /** the comment, or the scores that triggered it */
  text: string;
  /** the response itself */
  response: ResponseRecord;
}

/**
 * an issue before it has an id, club, date or status
 */
interface IssueDraft {
  /** which check */
  category: IssueCategoryKey;
  /** the event it is about */
  event: EventRecord;
  /** team the issue is about */
  team: string;
  /** text for the Details cell */
  details: string;
  /** the flagged responses, empty when there are none */
  responses: ResponseRecord[];
}

/**
 * one team averaging below the monitoring threshold at one tournament
 */
interface ClubBreach {
  /** club of the team */
  club: string;
  /** team that averaged below the threshold */
  team: string;
  /** the tournament */
  event: EventRecord;
  /** every score the team received there */
  responses: ResponseRecord[];
  /** the team's average there */
  average: number;
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
 * scores text with the response's comment quoted below it, if it has one
 *
 * @param text      the scores text
 * @param response  the response
 * @returns the text, with the comment on a new line
 */
function _withComment_(text: string, response: ResponseRecord): string {
  return response.comment === "" ? text : `${text}\n"${response.comment}"`;
}

/**
 * run the per-response checks over every response
 * pure, no google calls
 * skips events that are not included
 * min category and single low score are about the receiving team
 *   not in the policy, extra checks for the committee
 *   also run for international events, for awareness
 * the comment checks skip international events
 *   comments are not available and the scoring teams are not UKU teams
 *
 * @param responses  all responses
 * @param events     all events
 * @param settings   thresholds from the Issue Rules settings
 * @returns one hit per response per check it triggered
 */
function _responseIssueHits_(
  responses: ResponseRecord[],
  events: EventRecord[],
  settings: IssueSettings,
): IssueHit[] {
  const eventById = new Map(events.map((e) => [e.fileId, e]));
  const commentChecks = COMMENT_KEYWORD_CHECKS.map((category) => ({
    category,
    pattern: _keywordPattern_(settings.commentKeywords[category]),
  }));
  const hits: IssueHit[] = [];

  for (const r of responses) {
    const event = eventById.get(r.fileId);
    if (!event || !event.include) continue;

    const total = _responseTotal_(r);

    const minimums = SCORE_KEYS.filter((key) => r[key] === settings.categoryMinimum);
    if (minimums.length > 0) {
      hits.push({
        category: "categoryMinimum",
        event,
        team: r.receiver,
        other: r.scorer,
        text: _withComment_(minimums.map((key) => `${RESPONSE_HEADERS[key]} ${r[key]}`).join(", "), r),
        response: r,
      });
    }
    if (total < settings.singleLowScoreBelow) {
      hits.push({
        category: "singleLowScore",
        event,
        team: r.receiver,
        other: r.scorer,
        text: _withComment_(`Total ${total}`, r),
        response: r,
      });
    }

    if (event.international) continue;

    if (r.comment === "") {
      if (total > settings.commentTotalAbove || total < settings.commentTotalBelow) {
        hits.push({
          category: "totalWithoutComment",
          event,
          team: r.scorer,
          other: r.receiver,
          text: `Total ${total}`,
          response: r,
        });
      }
      const flagged = SCORE_KEYS.filter((key) => settings.commentCategoryScores.includes(r[key]));
      if (flagged.length > 0) {
        const scores = flagged.map((key) => `${RESPONSE_HEADERS[key]} ${r[key]}`).join(", ");
        hits.push({
          category: "categoryWithoutComment",
          event,
          team: r.scorer,
          other: r.receiver,
          text: scores,
          response: r,
        });
      }
    } else {
      for (const check of commentChecks) {
        if (!check.pattern || !check.pattern.test(r.comment)) continue;
        hits.push({
          category: check.category,
          event,
          team: r.receiver,
          other: r.scorer,
          text: r.comment,
          response: r,
        });
      }
    }
  }
  return hits;
}

/**
 * text for the Details cell of one per-response issue
 * a count, then one block per response
 *   comments are quoted in full, scores are listed
 *   "From" when the team received the response, "To" when the team gave it
 *
 * @param hits  every hit for this issue, all from the same check and event
 * @returns the details text
 */
function _issueDetails_(hits: IssueHit[]): string {
  const first = hits[0];
  const isComment = (COMMENT_KEYWORD_CHECKS as readonly string[]).includes(first.category);
  const count = `Number of ${isComment ? "comments" : "scores"}: ${hits.length}`;
  const blocks = hits.map((hit) => {
    const received = _teamKey_(hit.team) === _teamKey_(hit.response.receiver);
    const text = isComment ? `"${hit.text}"` : hit.text;
    return `${received ? "From" : "To"} ${hit.other}:\n${text}`;
  });
  return `${count}\n\n${blocks.join("\n\n")}`;
}

/**
 * group one event's responses by the team receiving them
 * pure, never modifies its arguments
 *
 * @param atEvent  responses from one event
 * @returns team key → responses received, in the order given
 */
function _receivedByTeam_(atEvent: ResponseRecord[]): Map<string, ResponseRecord[]> {
  const receivedBy: Map<string, ResponseRecord[]> = new Map();
  for (const r of atEvent) {
    const key = _teamKey_(r.receiver);
    receivedBy.set(key, [...(receivedBy.get(key) ?? []), r]);
  }
  return receivedBy;
}

/**
 * text for the Details cell of an average issue
 * the average, then every score received
 *
 * @param received  every score one team received at one event
 * @returns the details text
 */
function _averageDetails_(received: ResponseRecord[]): string {
  const totals = received.map(_responseTotal_);
  const average = totals.reduce((a, b) => a + b, 0) / totals.length;
  const lines = received.map((r) => `From ${r.scorer}:\nTotal ${_responseTotal_(r)}`);
  return `Average: ${average.toFixed(2)} from ${totals.length} scores\n\n${lines.join("\n\n")}`;
}

/**
 * the issue id for a check, event and team
 * readable, and the same every run so an issue is never added twice
 *
 * @param category  which check
 * @param event     the event
 * @param team      the team
 * @returns e.g. "LOW-AVERAGE | 2025-11-01 ELUXIR | Durham 1"
 */
function _issueId_(category: IssueCategoryKey, event: EventRecord, team: string): string {
  const date = event.date ? _isoDate_(event.date) : "no date";
  return `${ISSUE_CATEGORIES[category].code} | ${date} ${event.tournament} | ${team}`;
}

/**
 * group per-response hits into drafts, one per check per team per tournament
 * pure, never modifies its arguments
 *
 * @param hits  hits from the per-response checks
 * @returns drafts in the order they were first hit
 */
function _draftsFromHits_(hits: IssueHit[]): IssueDraft[] {
  const hitsById: Map<string, IssueHit[]> = new Map();
  for (const hit of hits) {
    const id = _issueId_(hit.category, hit.event, hit.team);
    hitsById.set(id, [...(hitsById.get(id) ?? []), hit]);
  }
  return [...hitsById.values()].map((grouped) => ({
    category: grouped[0].category,
    event: grouped[0].event,
    team: grouped[0].team,
    details: _issueDetails_(grouped),
    responses: grouped.map((hit) => hit.response),
  }));
}

/**
 * run the per-tournament checks for every team at every included event
 * pure, no google calls
 * low scores and low average also run for international events, for awareness
 * scores not submitted is skipped for international events
 *   the scoring teams there are not UKU teams
 *
 * @param responses  all responses, in date order
 * @param events     all events, in date order
 * @param settings   thresholds from the Issue Rules settings
 * @returns one draft per check per team per event that triggered
 */
function _teamEventIssueDrafts_(
  responses: ResponseRecord[],
  events: EventRecord[],
  settings: IssueSettings,
): IssueDraft[] {
  const drafts: IssueDraft[] = [];

  for (const event of events.filter((e) => e.include)) {
    const atEvent = responses.filter((r) => r.fileId === event.fileId);

    for (const received of _receivedByTeam_(atEvent).values()) {
      const team = received[0].receiver;
      const totals = received.map(_responseTotal_);

      const low = received.filter((r) => _responseTotal_(r) <= settings.lowScoreAtOrBelow);
      if (low.length >= settings.lowScoreCount) {
        const lines = low.map((r) => `From ${r.scorer}:\nTotal ${_responseTotal_(r)}`);
        drafts.push({
          category: "twoLowScores",
          event,
          team,
          details: `Number of scores: ${low.length}\n\n${lines.join("\n\n")}`,
          responses: low,
        });
      }

      const average = totals.reduce((a, b) => a + b, 0) / totals.length;
      if (average < settings.lowAverageBelow) {
        drafts.push({
          category: "lowAverage",
          event,
          team,
          details: _averageDetails_(received),
          responses: received,
        });
      }

      if (!event.international) {
        const key = _teamKey_(team);
        const submitted = atEvent.filter((r) => _teamKey_(r.scorer) === key);
        const missing: string[] = [];
        for (const opponent of new Set(received.map((r) => _teamKey_(r.scorer)))) {
          const from = received.filter((r) => _teamKey_(r.scorer) === opponent);
          const to = submitted.filter((r) => _teamKey_(r.receiver) === opponent).length;
          for (let i = to; i < from.length; i++) missing.push(from[0].scorer);
        }
        if (missing.length > 0) {
          const heading = `Received ${received.length}, submitted ${submitted.length}`;
          drafts.push({
            category: "notSubmitted",
            event,
            team,
            details: `${heading}\n\nMissing for:\n${missing.join("\n")}`,
            responses: [],
          });
        }
      }
    }
  }
  return drafts;
}

/**
 * one event's responses, indexed for the score columns
 */
interface EventScores {
  /** receiving team key → every response the team received there, in order */
  received: Map<string, ResponseRecord[]>;
  /** "scorer key|receiver key" → every response from that scorer to that receiver there, in order */
  pairs: Map<string, ResponseRecord[]>;
}

/**
 * index every response by event, receiving team and scorer → receiver pair
 * built once per refresh, so each issue looks up its scores instead of searching every response
 * pure, never modifies its arguments
 *
 * @param responses  all responses
 * @returns file id → that event's responses, indexed
 */
function _eventScores_(responses: ResponseRecord[]): Map<string, EventScores> {
  const byEvent: Map<string, EventScores> = new Map();
  for (const r of responses) {
    let scores = byEvent.get(r.fileId);
    if (!scores) {
      scores = { received: new Map(), pairs: new Map() };
      byEvent.set(r.fileId, scores);
    }
    const receiver = _teamKey_(r.receiver);
    const pair = `${_teamKey_(r.scorer)}|${receiver}`;
    const received = scores.received.get(receiver);
    if (received) received.push(r);
    else scores.received.set(receiver, [r]);
    const between = scores.pairs.get(pair);
    if (between) between.push(r);
    else scores.pairs.set(pair, [r]);
  }
  return byEvent;
}

/**
 * the Received, Given and Tournament Average cells for one draft
 * received and given have one value per flagged response, comma separated
 * tournament average has one value per team that received a flagged score
 * all blank when nothing was flagged
 * pure, never modifies its arguments
 *
 * received: total of the flagged response
 * given: total of the reply, the receiver scoring the scorer at the same tournament
 *   a pair that played twice is matched first to first, second to second
 * tournament average: average received at the tournament by each team that received a flagged score
 *   the flagged responses are left out, except for average issues where the full average is the point
 *
 * @param draft   the draft
 * @param scores  every event's responses, from _eventScores_
 * @returns the three cells
 */
function _issueScoreColumns_(
  draft: IssueDraft,
  scores: Map<string, EventScores>,
): { received: string; given: string; tournamentAverage: string } {
  if (draft.responses.length === 0) return { received: "", given: "", tournamentAverage: "" };

  const atEvent = scores.get(draft.event.fileId);
  const between = (scorer: string, receiver: string) =>
    atEvent?.pairs.get(`${_teamKey_(scorer)}|${_teamKey_(receiver)}`) ?? [];
  const isAverage = draft.category === "lowAverage" || draft.category === "monitoringBreach";
  const excluded: Set<ResponseRecord> = new Set(isAverage ? [] : draft.responses);

  const received = draft.responses.map((r) => String(_responseTotal_(r)));

  const given = draft.responses.map((r) => {
    const index = between(r.scorer, r.receiver).indexOf(r);
    const reply = between(r.receiver, r.scorer)[index];
    return reply ? String(_responseTotal_(reply)) : "–";
  });

  const receivers = [...new Set(draft.responses.map((r) => _teamKey_(r.receiver)))];
  const tournamentAverage = receivers.map((key) => {
    const totals = (atEvent?.received.get(key) ?? [])
      .filter((x) => !excluded.has(x))
      .map(_responseTotal_);
    return totals.length === 0 ? "–" : (totals.reduce((a, b) => a + b, 0) / totals.length).toFixed(2);
  });

  return { received: received.join(", "), given: given.join(", "), tournamentAverage: tournamentAverage.join(", ") };
}

/**
 * turn drafts into Issues rows
 * pure, never modifies its arguments
 * INTERNATIONAL is put above the details for international tournaments
 *
 * @param drafts     drafts from every check
 * @param responses  all responses, for the score columns
 * @param clubOf     team key → club, from the Teams tab
 * @param today      date to record as the Issue Date
 * @param settings   thresholds, for the Issue Category text
 * @returns issues in the order given
 */
function _issueRecords_(
  drafts: IssueDraft[],
  responses: ResponseRecord[],
  clubOf: Map<string, string>,
  today: Date,
  settings: IssueSettings,
): IssueRecord[] {
  const scores = _eventScores_(responses);
  return drafts.map((draft) => ({
    issueId: _issueId_(draft.category, draft.event, draft.team),
    dateCreated: today,
    club: clubOf.get(_teamKey_(draft.team)) ?? "",
    category: ISSUE_CATEGORIES[draft.category].label(settings),
    team: draft.team,
    tournaments: draft.event.tournament,
    tournamentDate: draft.event.date,
    details: `${draft.event.international ? "INTERNATIONAL\n\n" : ""}${draft.details}`,
    ..._issueScoreColumns_(draft, scores),
    status: ISSUE_STATUSES[0],
    committeeMember: "",
    notes: "",
  }));
}

/**
 * find every monitoring breach, grouped by club
 * pure, no google calls
 * a breach is a team averaging below the threshold at a tournament
 *   two teams from one club at the same tournament are two breaches
 * skips events that are not included, and international events
 *
 * @param responses  all responses
 * @param events     all events, in date order
 * @param clubOf     team key → club, from the Teams tab
 * @param settings   thresholds from the Issue Rules settings
 * @returns club → its breaches in date order, teams without a club are skipped
 */
function _clubBreaches_(
  responses: ResponseRecord[],
  events: EventRecord[],
  clubOf: Map<string, string>,
  settings: IssueSettings,
): Map<string, ClubBreach[]> {
  const breachesByClub: Map<string, ClubBreach[]> = new Map();
  for (const event of events.filter((e) => e.include && !e.international)) {
    const atEvent = responses.filter((r) => r.fileId === event.fileId);
    for (const [key, received] of _receivedByTeam_(atEvent)) {
      const club = clubOf.get(key) ?? "";
      const average = received.map(_responseTotal_).reduce((a, b) => a + b, 0) / received.length;
      if (club === "" || average >= settings.monitoringAverageBelow) continue;
      const breach = { club, team: received[0].receiver, event, responses: received, average };
      breachesByClub.set(club, [...(breachesByClub.get(club) ?? []), breach]);
    }
  }
  return breachesByClub;
}

/**
 * Issues rows for the club monitoring list
 * pure, never modifies its arguments
 * a club goes on the list at its second breach, one MONITORING row lists those breaches
 * every later breach gets its own MONITORING-BREACH row
 * a hub covers one season, so the MONITORING id is just the club
 *
 * @param breachesByClub  from _clubBreaches_
 * @param responses       all responses, for the score columns of breach rows
 * @param clubOf          team key → club, from the Teams tab
 * @param today           date to record as the Issue Date
 * @param settings        thresholds and which checks are enabled
 * @returns MONITORING rows, then MONITORING-BREACH rows
 */
function _monitoringIssues_(
  breachesByClub: Map<string, ClubBreach[]>,
  responses: ResponseRecord[],
  clubOf: Map<string, string>,
  today: Date,
  settings: IssueSettings,
): IssueRecord[] {
  if (!settings.enabled.monitoring) return [];
  const needed = settings.monitoringBreaches;
  const dateOf = (b: ClubBreach) => (b.event.date ? _isoDate_(b.event.date) : "no date");
  const listed: IssueRecord[] = [];
  const later: IssueDraft[] = [];

  for (const [club, breaches] of breachesByClub) {
    if (breaches.length < needed) continue;
    const first = breaches.slice(0, needed);
    const blocks = first.map((b) =>
      `${b.event.tournament} (${dateOf(b)}):\n${b.team} averaged ${
        b.average.toFixed(2)
      } from ${b.responses.length} scores`
    );
    listed.push({
      issueId: `${ISSUE_CATEGORIES.monitoring.code} | ${club}`,
      dateCreated: today,
      club,
      category: ISSUE_CATEGORIES.monitoring.label(settings),
      team: first.map((b) => b.team).join(", "),
      tournaments: first.map((b) => b.event.tournament).join(", "),
      tournamentDate: first.map(dateOf).join(", "),
      details: `Number of breaches: ${first.length}\n\n${blocks.join("\n\n")}`,
      received: "",
      given: "",
      tournamentAverage: first.map((b) => b.average.toFixed(2)).join(", "),
      status: ISSUE_STATUSES[0],
      committeeMember: "",
      notes: "",
    });
    for (const b of settings.enabled.monitoringBreach ? breaches.slice(needed) : []) {
      later.push({
        category: "monitoringBreach",
        event: b.event,
        team: b.team,
        details: _averageDetails_(b.responses),
        responses: b.responses,
      });
    }
  }
  return [...listed, ..._issueRecords_(later, responses, clubOf, today, settings)];
}

/**
 * @returns the Issues tab, created on first use
 */
function _getIssuesSheet_(): GoogleAppsScript.Spreadsheet.Sheet {
  const headers: readonly string[] = Object.values(ISSUE_HEADERS);
  return _getOrCreateSheet_(SpreadsheetApp.getActiveSpreadsheet(), ISSUES_SHEET, headers, ISSUES_INFO);
}

/**
 * add issues that are not already on the Issues tab, at the bottom
 * existing rows are never read back or rewritten, only their Issue IDs are checked
 *
 * @param issues  every issue that should exist
 * @returns how many new issues were added
 */
function _appendIssues_(issues: IssueRecord[]): number {
  const sheet = _getIssuesSheet_();
  const lastRow = sheet.getLastRow();
  const idColumn = ISSUE_KEYS.indexOf("issueId") + 1;
  const existingIds = new Set(
    lastRow < DATA_ROW
      ? []
      : sheet.getRange(DATA_ROW, idColumn, lastRow - HEADER_ROW, 1).getValues().map((r) => String(r[0])),
  );

  const added = issues.filter((issue) => !existingIds.has(issue.issueId));
  if (added.length === 0) return 0;

  const firstNew = Math.max(lastRow, HEADER_ROW) + 1;
  _fitSheet_(sheet, firstNew + added.length - 1, ISSUE_KEYS.length);
  sheet.getRange(firstNew, 1, added.length, ISSUE_KEYS.length).setValues(
    added.map((issue) => ISSUE_KEYS.map((key) => issue[key] ?? "")),
  );

  const column = (key: keyof typeof ISSUE_HEADERS) => ISSUE_KEYS.indexOf(key) + 1;
  const statusRule = SpreadsheetApp.newDataValidation().requireValueInList([...ISSUE_STATUSES], true).build();
  sheet.getRange(firstNew, column("status"), added.length, 1).setDataValidation(statusRule);
  for (const key of ["dateCreated", "tournamentDate"] as const) {
    sheet.getRange(firstNew, column(key), added.length, 1).setNumberFormat("yyyy-mm-dd");
  }
  _applyFilter_(sheet, firstNew + added.length - 1 - HEADER_ROW, ISSUE_KEYS.length);
  return added.length;
}

/**
 * run every check on what is already in the spreadsheet and add the new issues, no files are read
 * use after Refresh Results, or after changing Include, International, a club override or an issue rule
 * tabs may have been sorted by hand, so events and responses are put back in date order first
 * reads the clubs from the Teams tab and the thresholds from the Issue Rules settings, disabled checks add nothing
 *
 * @returns a one-line summary
 */
function _refreshIssues_(): string {
  const { events, responses, settings, clubOf } = _timed_("read for issues", () => {
    const events = _sortEvents_(_readEvents_(_getEventsSheet_()));
    const responses = _sortResponses_(_readResponses_(_getResponsesSheet_()), events);
    return { events, responses, settings: _readIssueSettings_(), clubOf: _readTeamClubs_() };
  });
  if (responses.length === 0) return "No results found: run Refresh Results first";

  const issues = _timed_("check issues", () => {
    const drafts = [
      ..._draftsFromHits_(_responseIssueHits_(responses, events, settings)),
      ..._teamEventIssueDrafts_(responses, events, settings),
    ].filter((draft) => settings.enabled[draft.category]);
    const today = new Date();
    const breaches = _clubBreaches_(responses, events, clubOf, settings);
    return [
      ..._issueRecords_(drafts, responses, clubOf, today, settings),
      ..._monitoringIssues_(breaches, responses, clubOf, today, settings),
    ];
  });
  const added = _timed_("write issues", () => _appendIssues_(issues));
  return `${added} new issue(s)`;
}
