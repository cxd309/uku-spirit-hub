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
  + "New issues are added on each import, existing rows are never changed, so sort and filter freely";

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
 * @property {Date|null} tournamentDate     tournament date, null if unknown
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
 * skips events that are not included, and international events
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
    if (!event || !event.include || event.international) continue;

    const total = _responseTotal_(r);

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
 *
 * @param {IssueHit[]} hits  every hit for this issue, all from the same check and event
 * @returns {string} the details text
 */
function _issueDetails_(hits) {
  const first = hits[0];
  const isComment = first.category === "dangerousPlay";
  const count = `Number of ${isComment ? "comments" : "scores"}: ${hits.length}`;
  const blocks = hits.map((hit) => isComment ? `From ${hit.other}:\n"${hit.text}"` : `To ${hit.other}:\n${hit.text}`);
  return `${count}\n\n${blocks.join("\n\n")}`;
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

    /** @type {Map<string, ResponseRecord[]>} */
    const receivedBy = new Map();
    for (const r of atEvent) {
      const key = _teamKey_(r.receiver);
      receivedBy.set(key, [...(receivedBy.get(key) ?? []), r]);
    }

    for (const received of receivedBy.values()) {
      const team = received[0].receiver;
      const totals = received.map(_responseTotal_);
      const scoreLines = received.map((r) => `From ${r.scorer}:\nTotal ${_responseTotal_(r)}`);

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
        const heading = `Average: ${average.toFixed(2)} from ${totals.length} scores`;
        drafts.push({
          category: "lowAverage",
          event,
          team,
          details: `${heading}\n\n${scoreLines.join("\n\n")}`,
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
 *   the flagged responses are left out, except for low average where the full average is the point
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
  const excluded = draft.category === "lowAverage" ? [] : draft.responses;

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
