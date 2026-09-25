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
  + "- Status: NEW, Contacted or Closed\n"
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
    club: "Club",
    category: "Issue Category",
    team: "Team",
    tournaments: "Tournament(s)",
    details: "Details",
    dateCreated: "Date Created",
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
});

/**
 * one row of the Issues tab
 *
 * @typedef {Object} IssueRecord
 * @property {string}  issueId          check, tournament and team, unique per issue
 * @property {string}  club             club the issue is about
 * @property {string}  category         label of the check
 * @property {string}  team             team the issue is about
 * @property {string}  tournaments      tournament name(s)
 * @property {string}  details          one line per response that triggered it
 * @property {Date}    dateCreated      when the issue was first added
 * @property {string}  status           one of ISSUE_STATUSES
 * @property {string}  committeeMember  who is handling it
 * @property {string}  notes            free text
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
 */

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

    const total = SCORE_KEYS.reduce((sum, key) => sum + r[key], 0);

    if (r.comment === "") {
      if (total > ISSUE_THRESHOLDS.commentTotalAbove || total < ISSUE_THRESHOLDS.commentTotalBelow) {
        hits.push({
          category: "totalWithoutComment",
          event,
          team: r.scorer,
          other: r.receiver,
          text: `Total ${total}`,
        });
      }
      const flagged = SCORE_KEYS.filter((key) => ISSUE_THRESHOLDS.commentCategoryScores.includes(r[key]));
      if (flagged.length > 0) {
        const scores = flagged.map((key) => `${RESPONSE_HEADERS[key]} ${r[key]}`).join(", ");
        hits.push({ category: "categoryWithoutComment", event, team: r.scorer, other: r.receiver, text: scores });
      }
    } else if (keywords.test(r.comment)) {
      hits.push({ category: "dangerousPlay", event, team: r.receiver, other: r.scorer, text: r.comment });
    }
  }
  return hits;
}

/**
 * text for the Details cell of one issue
 * INTERNATIONAL first when the tournament is international
 * then a count, then one block per response
 *   comments are quoted in full, scores are listed
 *
 * @param {IssueHit[]} hits  every hit for this issue, all from the same check and event
 * @returns {string} the details text
 */
function _issueDetails_(hits) {
  const first = hits[0];
  const isComment = first.category === "dangerousPlay";
  const header = first.event.international ? "INTERNATIONAL\n\n" : "";
  const count = `Number of ${isComment ? "comments" : "scores"}: ${hits.length}`;
  const blocks = hits.map((hit) => isComment ? `From ${hit.other}:\n"${hit.text}"` : `To ${hit.other}:\n${hit.text}`);
  return `${header}${count}\n\n${blocks.join("\n\n")}`;
}

/**
 * group hits into issues, one per check per team per tournament
 * pure, never modifies its arguments
 *
 * @param {IssueHit[]}          hits    hits from the checks
 * @param {Map<string, string>} clubOf  team key → club, from the Teams tab
 * @param {Date}                today   date to record as Date Created
 * @returns {IssueRecord[]} issues in the order they were first hit
 */
function _groupIssueHits_(hits, clubOf, today) {
  /** @type {Map<string, IssueHit[]>} */
  const hitsById = new Map();
  for (const hit of hits) {
    const eventLabel = `${hit.event.date ? _isoDate_(hit.event.date) : "no date"} ${hit.event.tournament}`;
    const issueId = `${ISSUE_CATEGORIES[hit.category].code} | ${eventLabel} | ${hit.team}`;
    hitsById.set(issueId, [...(hitsById.get(issueId) ?? []), hit]);
  }

  return [...hitsById].map(([issueId, grouped]) => ({
    issueId: issueId,
    club: clubOf.get(_teamKey_(grouped[0].team)) ?? "",
    category: ISSUE_CATEGORIES[grouped[0].category].label,
    team: grouped[0].team,
    tournaments: grouped[0].event.tournament,
    details: _issueDetails_(grouped),
    dateCreated: today,
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
    added.map((issue) => ISSUE_KEYS.map((key) => issue[key])),
  );

  /** @param {keyof typeof ISSUE_HEADERS} key */
  const column = (key) => ISSUE_KEYS.indexOf(key) + 1;
  const statusRule = SpreadsheetApp.newDataValidation().requireValueInList([...ISSUE_STATUSES], true).build();
  sheet.getRange(firstNew, column("status"), added.length, 1).setDataValidation(statusRule);
  sheet.getRange(firstNew, column("dateCreated"), added.length, 1).setNumberFormat("yyyy-mm-dd");
  _applyFilter_(sheet, firstNew + added.length - 1 - HEADER_ROW, ISSUE_KEYS.length);
  return added.length;
}
