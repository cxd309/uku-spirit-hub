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
  + "New issues are added by Refresh Results and Refresh Issues, existing rows are never changed, so sort and filter freely";

/**
 * issues table columns, in order
 * keys are the names used in code
 * values are the header text
 */
const ISSUE_HEADERS = Object.freeze(
  /** @type {const} */ ({
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
  }),
);

/**
 * keys of the Issues tab in column order
 */
const ISSUE_KEYS = /** @type {(keyof typeof ISSUE_HEADERS)[]} */ (Object.keys(ISSUE_HEADERS));

/**
 * values of the Status column, first is given to new issues
 */
const ISSUE_STATUSES = Object.freeze(["NEW", "IN PROGRESS", "CLOSED"]);

/**
 * policy thresholds used by the checks
 * kept in one place so they can move to the Config tab later
 */
const ISSUE_THRESHOLDS = Object.freeze({
  commentTotalAbove: 14,
  commentTotalBelow: 6,
  commentCategoryScores: Object.freeze([0, 4]),
  lowScoreAtOrBelow: 6,
  lowScoreCount: 2,
  lowAverageBelow: 8,
  monitoringAverageBelow: 9,
  monitoringBreaches: 2,
  categoryMinimum: 0,
  singleLowScoreBelow: 6,
});

/**
 * words in a comment that flag possible dangerous play
 * matched as whole words, case-insensitive
 */
const DANGEROUS_PLAY_KEYWORDS = Object.freeze([
  "dangerous",
  "danger",
  "reckless",
  "unsafe",
]);

/**
 * every issue check
 * code goes into the Issue ID, label into the Issue Category column
 */
const ISSUE_CATEGORIES = Object.freeze({
  totalWithoutComment: {
    code: "TOTAL-NO-COMMENT",
    label:
      `Total above ${ISSUE_THRESHOLDS.commentTotalAbove} or below ${ISSUE_THRESHOLDS.commentTotalBelow} without comment`,
  },
  categoryWithoutComment: {
    code: "CATEGORY-NO-COMMENT",
    label: `Category scored ${ISSUE_THRESHOLDS.commentCategoryScores.join(" or ")} without comment`,
  },
  dangerousPlay: {
    code: "DANGEROUS-PLAY",
    label: "Dangerous play mentioned",
  },
  notSubmitted: {
    code: "NOT-SUBMITTED",
    label: "Spirit scores not submitted",
  },
  twoLowScores: {
    code: "LOW-SCORES",
    label:
      `${ISSUE_THRESHOLDS.lowScoreCount} or more scores of ${ISSUE_THRESHOLDS.lowScoreAtOrBelow} or below at a tournament`,
  },
  lowAverage: {
    code: "LOW-AVERAGE",
    label: `Average score below ${ISSUE_THRESHOLDS.lowAverageBelow} at a tournament`,
  },
  categoryMinimum: {
    code: "MIN-CATEGORY",
    label: `Received ${ISSUE_THRESHOLDS.categoryMinimum} in a category`,
  },
  singleLowScore: {
    code: "SINGLE-LOW-SCORE",
    label: `Received a score below ${ISSUE_THRESHOLDS.singleLowScoreBelow}`,
  },
  monitoring: {
    code: "MONITORING",
    label:
      `${ISSUE_THRESHOLDS.monitoringBreaches} team averages below ${ISSUE_THRESHOLDS.monitoringAverageBelow} in the season`,
  },
  monitoringBreach: {
    code: "MONITORING-BREACH",
    label: `Averaged below ${ISSUE_THRESHOLDS.monitoringAverageBelow} again while on monitoring`,
  },
});

/**
 * one row of the Issues tab
 *
 * @typedef {Object} IssueRecord
 * @property {string}    issueId            check, tournament and team, unique per issue
 * @property {Date}      dateCreated        when the issue was first added
 * @property {string}    club               club the issue is about
 * @property {string}    category           label of the check
 * @property {string}    team               team the issue is about
 * @property {string}    tournaments        tournament name(s)
 * @property {Date|string|null} tournamentDate  tournament date, null if unknown, comma separated text for several
 * @property {string}    details            one line per response that triggered it
 * @property {string}    received           total of each flagged response, comma separated
 * @property {string}    given              total of the reply to each flagged response, comma separated
 * @property {string}    tournamentAverage  average received at the tournament by each team that received a flagged score
 * @property {string}    status             one of ISSUE_STATUSES
 * @property {string}    committeeMember    who is handling it
 * @property {string}    notes              free text
 */

/**
 * one response that triggered a check, before grouping into issues
 *
 * @typedef {Object} IssueHit
 * @property {keyof typeof ISSUE_CATEGORIES} category  which check
 * @property {EventRecord}                   event     the event the response is from
 * @property {string}                        team      team the issue is about
 * @property {string}                        other     the other team in the response
 * @property {string}                        text      the comment, or the scores that triggered it
 * @property {ResponseRecord}                response  the response itself
 */

/**
 * an issue before it has an id, club, date or status
 *
 * @typedef {Object} IssueDraft
 * @property {keyof typeof ISSUE_CATEGORIES} category  which check
 * @property {EventRecord}                   event     the event it is about
 * @property {string}                        team      team the issue is about
 * @property {string}                        details   text for the Details cell
 * @property {ResponseRecord[]}              responses the flagged responses, empty when there are none
 */

/**
 * one team averaging below the monitoring threshold at one tournament
 *
 * @typedef {Object} ClubBreach
 * @property {string}           club       club of the team
 * @property {string}           team       team that averaged below the threshold
 * @property {EventRecord}      event      the tournament
 * @property {ResponseRecord[]} responses  every score the team received there
 * @property {number}           average    the team's average there
 */

/**
 * total of the five category scores in a response
 *
 * @param {ResponseRecord} response  the response
 * @returns {number} the total
 */
function _responseTotal_(response) {
  return SCORE_KEYS.reduce((sum, key) => sum + response[key], 0);
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
 * @param {ResponseRecord[]} responses  all responses
 * @param {EventRecord[]}    events     all events
 * @returns {IssueHit[]} one hit per response per check it triggered
 */
function _responseIssueHits_(responses, events) {
  const eventById = new Map(events.map((e) => [e.fileId, e]));
  const keywords = new RegExp(`\\b(${DANGEROUS_PLAY_KEYWORDS.join("|")})\\b`, "i");
  /** @type {IssueHit[]} */
  const hits = [];

  for (const r of responses) {
    const event = eventById.get(r.fileId);
    if (!event || !event.include) continue;

    const total = _responseTotal_(r);

    const minimums = SCORE_KEYS.filter((key) => r[key] === ISSUE_THRESHOLDS.categoryMinimum);
    if (minimums.length > 0) {
      hits.push({
        category: "categoryMinimum",
        event,
        team: r.receiver,
        other: r.scorer,
        text: minimums.map((key) => `${RESPONSE_HEADERS[key]} ${r[key]}`).join(", "),
        response: r,
      });
    }
    if (total < ISSUE_THRESHOLDS.singleLowScoreBelow) {
      hits.push({ category: "singleLowScore", event, team: r.receiver, other: r.scorer, text: `Total ${total}`, response: r });
    }

    if (event.international) continue;

    if (r.comment === "") {
      if (total > ISSUE_THRESHOLDS.commentTotalAbove || total < ISSUE_THRESHOLDS.commentTotalBelow) {
        hits.push({
          category: "totalWithoutComment",
          event,
          team: r.scorer,
          other: r.receiver,
          text: `Total ${total}`,
          response: r,
        });
      }
      const flagged = SCORE_KEYS.filter((key) => ISSUE_THRESHOLDS.commentCategoryScores.includes(r[key]));
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
    } else if (keywords.test(r.comment)) {
      hits.push({
        category: "dangerousPlay",
        event,
        team: r.receiver,
        other: r.scorer,
        text: r.comment,
        response: r,
      });
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
 * @param {IssueHit[]} hits  every hit for this issue, all from the same check and event
 * @returns {string} the details text
 */
function _issueDetails_(hits) {
  const first = hits[0];
  const isComment = first.category === "dangerousPlay";
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
 * @param {ResponseRecord[]} atEvent  responses from one event
 * @returns {Map<string, ResponseRecord[]>} team key → responses received, in the order given
 */
function _receivedByTeam_(atEvent) {
  /** @type {Map<string, ResponseRecord[]>} */
  const receivedBy = new Map();
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
 * @param {ResponseRecord[]} received  every score one team received at one event
 * @returns {string} the details text
 */
function _averageDetails_(received) {
  const totals = received.map(_responseTotal_);
  const average = totals.reduce((a, b) => a + b, 0) / totals.length;
  const lines = received.map((r) => `From ${r.scorer}:\nTotal ${_responseTotal_(r)}`);
  return `Average: ${average.toFixed(2)} from ${totals.length} scores\n\n${lines.join("\n\n")}`;
}

/**
 * the issue id for a check, event and team
 * readable, and the same every run so an issue is never added twice
 *
 * @param {keyof typeof ISSUE_CATEGORIES} category  which check
 * @param {EventRecord}                   event     the event
 * @param {string}                        team      the team
 * @returns {string} e.g. "LOW-AVERAGE | 2025-11-01 ELUXIR | Durham 1"
 */
function _issueId_(category, event, team) {
  const date = event.date ? _isoDate_(event.date) : "no date";
  return `${ISSUE_CATEGORIES[category].code} | ${date} ${event.tournament} | ${team}`;
}

/**
 * group per-response hits into drafts, one per check per team per tournament
 * pure, never modifies its arguments
 *
 * @param {IssueHit[]} hits  hits from the per-response checks
 * @returns {IssueDraft[]} drafts in the order they were first hit
 */
function _draftsFromHits_(hits) {
  /** @type {Map<string, IssueHit[]>} */
  const hitsById = new Map();
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
 * @param {ResponseRecord[]} responses  all responses, in date order
 * @param {EventRecord[]}    events     all events, in date order
 * @returns {IssueDraft[]} one draft per check per team per event that triggered
 */
function _teamEventIssueDrafts_(responses, events) {
  /** @type {IssueDraft[]} */
  const drafts = [];

  for (const event of events.filter((e) => e.include)) {
    const atEvent = responses.filter((r) => r.fileId === event.fileId);

    for (const received of _receivedByTeam_(atEvent).values()) {
      const team = received[0].receiver;
      const totals = received.map(_responseTotal_);

      const low = received.filter((r) => _responseTotal_(r) <= ISSUE_THRESHOLDS.lowScoreAtOrBelow);
      if (low.length >= ISSUE_THRESHOLDS.lowScoreCount) {
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
      if (average < ISSUE_THRESHOLDS.lowAverageBelow) {
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
        /** @type {string[]} */
        const missing = [];
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
 * @param {IssueDraft}       draft      the draft
 * @param {ResponseRecord[]} responses  all responses
 * @returns {{received: string, given: string, tournamentAverage: string}} the three cells
 */
function _issueScoreColumns_(draft, responses) {
  if (draft.responses.length === 0) return { received: "", given: "", tournamentAverage: "" };

  const atEvent = responses.filter((r) => r.fileId === draft.event.fileId);
  /**
   * @param {ResponseRecord} r
   * @param {string} scorer
   * @param {string} receiver
   */
  const isPair = (r, scorer, receiver) =>
    _teamKey_(r.scorer) === _teamKey_(scorer) && _teamKey_(r.receiver) === _teamKey_(receiver);
  const isAverage = draft.category === "lowAverage" || draft.category === "monitoringBreach";
  const excluded = isAverage ? [] : draft.responses;

  const received = draft.responses.map((r) => String(_responseTotal_(r)));

  const given = draft.responses.map((r) => {
    const index = atEvent.filter((x) => isPair(x, r.scorer, r.receiver)).indexOf(r);
    const reply = atEvent.filter((x) => isPair(x, r.receiver, r.scorer))[index];
    return reply ? String(_responseTotal_(reply)) : "–";
  });

  const receivers = [...new Set(draft.responses.map((r) => _teamKey_(r.receiver)))];
  const tournamentAverage = receivers.map((key) => {
    const totals = atEvent
      .filter((x) => _teamKey_(x.receiver) === key && !excluded.includes(x))
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
 * @param {IssueDraft[]}        drafts     drafts from every check
 * @param {ResponseRecord[]}    responses  all responses, for the score columns
 * @param {Map<string, string>} clubOf     team key → club, from the Teams tab
 * @param {Date}                today      date to record as the Issue Date
 * @returns {IssueRecord[]} issues in the order given
 */
function _issueRecords_(drafts, responses, clubOf, today) {
  return drafts.map((draft) => ({
    issueId: _issueId_(draft.category, draft.event, draft.team),
    dateCreated: today,
    club: clubOf.get(_teamKey_(draft.team)) ?? "",
    category: ISSUE_CATEGORIES[draft.category].label,
    team: draft.team,
    tournaments: draft.event.tournament,
    tournamentDate: draft.event.date,
    details: `${draft.event.international ? "INTERNATIONAL\n\n" : ""}${draft.details}`,
    ..._issueScoreColumns_(draft, responses),
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
 * @param {ResponseRecord[]}    responses  all responses
 * @param {EventRecord[]}       events     all events, in date order
 * @param {Map<string, string>} clubOf     team key → club, from the Teams tab
 * @returns {Map<string, ClubBreach[]>} club → its breaches in date order, teams without a club are skipped
 */
function _clubBreaches_(responses, events, clubOf) {
  /** @type {Map<string, ClubBreach[]>} */
  const breachesByClub = new Map();
  for (const event of events.filter((e) => e.include && !e.international)) {
    const atEvent = responses.filter((r) => r.fileId === event.fileId);
    for (const [key, received] of _receivedByTeam_(atEvent)) {
      const club = clubOf.get(key) ?? "";
      const average = received.map(_responseTotal_).reduce((a, b) => a + b, 0) / received.length;
      if (club === "" || average >= ISSUE_THRESHOLDS.monitoringAverageBelow) continue;
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
 * @param {Map<string, ClubBreach[]>} breachesByClub  from _clubBreaches_
 * @param {ResponseRecord[]}          responses       all responses, for the score columns of breach rows
 * @param {Map<string, string>}       clubOf          team key → club, from the Teams tab
 * @param {Date}                      today           date to record as the Issue Date
 * @returns {IssueRecord[]} MONITORING rows, then MONITORING-BREACH rows
 */
function _monitoringIssues_(breachesByClub, responses, clubOf, today) {
  const needed = ISSUE_THRESHOLDS.monitoringBreaches;
  /** @param {ClubBreach} b */
  const dateOf = (b) => (b.event.date ? _isoDate_(b.event.date) : "no date");
  /** @type {IssueRecord[]} */
  const listed = [];
  /** @type {IssueDraft[]} */
  const later = [];

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
      category: ISSUE_CATEGORIES.monitoring.label,
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
    for (const b of breaches.slice(needed)) {
      later.push({
        category: "monitoringBreach",
        event: b.event,
        team: b.team,
        details: _averageDetails_(b.responses),
        responses: b.responses,
      });
    }
  }
  return [...listed, ..._issueRecords_(later, responses, clubOf, today)];
}

/**
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} the Issues tab, created on first use
 */
function _getIssuesSheet_() {
  /** @type {readonly string[]} */
  const headers = Object.values(ISSUE_HEADERS);
  return _getOrCreateSheet_(SpreadsheetApp.getActiveSpreadsheet(), ISSUES_SHEET, headers, ISSUES_INFO).sheet;
}

/**
 * add issues that are not already on the Issues tab, at the bottom
 * existing rows are never read back or rewritten, only their Issue IDs are checked
 *
 * @param {IssueRecord[]} issues  every issue that should exist
 * @returns {number} how many new issues were added
 */
function _appendIssues_(issues) {
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

  /** @param {keyof typeof ISSUE_HEADERS} key */
  const column = (key) => ISSUE_KEYS.indexOf(key) + 1;
  const statusRule = SpreadsheetApp.newDataValidation().requireValueInList([...ISSUE_STATUSES], true).build();
  sheet.getRange(firstNew, column("status"), added.length, 1).setDataValidation(statusRule);
  for (const key of /** @type {const} */ (["dateCreated", "tournamentDate"])) {
    sheet.getRange(firstNew, column(key), added.length, 1).setNumberFormat("yyyy-mm-dd");
  }
  _applyFilter_(sheet, firstNew + added.length - 1 - HEADER_ROW, ISSUE_KEYS.length);
  return added.length;
}

/**
 * run every check and add the issues that are not already on the Issues tab
 * reads the clubs from the Teams tab, so Teams must be up to date first
 *
 * @param {ResponseRecord[]} responses  all responses, in date order
 * @param {EventRecord[]}    events     all events, in date order
 * @returns {number} how many new issues were added
 */
function _appendNewIssues_(responses, events) {
  const drafts = [
    ..._draftsFromHits_(_responseIssueHits_(responses, events)),
    ..._teamEventIssueDrafts_(responses, events),
  ];
  const clubOf = _readTeamClubs_();
  const today = new Date();
  return _appendIssues_([
    ..._issueRecords_(drafts, responses, clubOf, today),
    ..._monitoringIssues_(_clubBreaches_(responses, events, clubOf), responses, clubOf, today),
  ]);
}

/**
 * run every check on what is already in the spreadsheet, no files are read
 * use after changing Include, International, a club override or a threshold
 * tabs may have been sorted by hand, so events and responses are put back in date order first
 *
 * @returns {string} a one-line summary
 */
function _refreshIssues_() {
  const events = _sortEvents_(_readEvents_(_getEventsSheet_()));
  const responses = _sortResponses_(_readResponses_(_getResponsesSheet_()), events);
  if (responses.length === 0) return "No results found: run Refresh Results first";
  return `${_appendNewIssues_(responses, events)} new issue(s)`;
}
