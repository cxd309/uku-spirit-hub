/**
 * header row of the Issue Rules table
 */
const ISSUE_RULE_HEADERS = Object.freeze(
  /** @type {const} */ (["Rule", "Enabled", "Value 1", "Value 2", "Description"]),
);

/**
 * the row written for each check when it is missing from the section
 * value 1 and value 2 mean different things per check, the description says what
 * lists are written as comma separated text
 *
 * @type {Readonly<Record<IssueCategoryKey, {value1: number|string, value2: number|string, description: string}>>}
 */
const DEFAULT_ISSUE_RULES = Object.freeze({
  totalWithoutComment: {
    value1: 14,
    value2: 6,
    description: "Policy: a single score total above Value 1 or below Value 2 needs a comment",
  },
  categoryWithoutComment: {
    value1: "0, 4",
    value2: "",
    description: "Policy: a category scored any of Value 1 needs a comment",
  },
  dangerousPlay: {
    value1: "dangerous, danger, reckless, unsafe",
    value2: "",
    description: "Policy: a comment contains any of the words in Value 1",
  },
  notSubmitted: {
    value1: "",
    value2: "",
    description: "Policy: a team gave fewer scores to an opponent than it received from them",
  },
  twoLowScores: {
    value1: 6,
    value2: 2,
    description: "Policy: Value 2 or more scores of Value 1 or below at a single tournament",
  },
  lowAverage: {
    value1: 8,
    value2: "",
    description: "Policy: an average below Value 1 at a single tournament tournament",
  },
  categoryMinimum: {
    value1: 0,
    value2: "",
    description: "Extra: received Value 1 in any category",
  },
  singleLowScore: {
    value1: 6,
    value2: "",
    description: "Extra: received a single score total below Value 1",
  },
  monitoring: {
    value1: 9,
    value2: 2,
    description:
      "Policy: a club's teams average below Value 1, Value 2 times in the season, the club goes on monitoring.",
  },
  monitoringBreach: {
    value1: "",
    value2: "",
    description: "Another average below the MONITORING Value 1 once a club is on monitoring. Needs MONITORING enabled",
  },
});

/**
 * thresholds and switches for every issue check, read from the Issue Rules settings
 *
 * @typedef {Object} IssueSettings
 * @property {Record<IssueCategoryKey, boolean>} enabled  which checks add issues
 * @property {number}   commentTotalAbove       total above this needs a comment
 * @property {number}   commentTotalBelow       total below this needs a comment
 * @property {number[]} commentCategoryScores   category scores that need a comment
 * @property {string[]} dangerousPlayKeywords   words that flag a comment
 * @property {number}   lowScoreAtOrBelow       a total at or below this is a low score
 * @property {number}   lowScoreCount           this many low scores at a tournament is an issue
 * @property {number}   lowAverageBelow         an average below this at a tournament is an issue
 * @property {number}   categoryMinimum         receiving this in a category is an issue
 * @property {number}   singleLowScoreBelow     receiving a total below this is an issue
 * @property {number}   monitoringAverageBelow  an average below this is a monitoring breach
 * @property {number}   monitoringBreaches      this many breaches puts a club on monitoring
 */

/**
 * every check key, in the order of ISSUE_CATEGORIES
 *
 * @returns {IssueCategoryKey[]} the keys
 */
function _issueCategoryKeys_() {
  return /** @type {IssueCategoryKey[]} */ (Object.keys(ISSUE_CATEGORIES));
}

/**
 * the Issue Rules section of the Settings tab
 * keyed by rule code, so a missing rule is added back with its defaults
 * and checks added in later versions appear on their own
 *
 * @type {SettingsSection}
 */
const ISSUE_RULES_SECTION = Object.freeze({
  title: "Issue Rules",
  description: "Settings for each issue check, used by Refresh Issues. "
    + "Untick Enabled to stop adding new issues of that type. "
    + "Value 1 and Value 2 are thresholds, the description says what each means, lists are comma separated. "
    + "Do not change the Rule column. Issues already on the Issues tab are never changed",
  headers: ISSUE_RULE_HEADERS,
  keyed: true,
  defaults: () =>
    _issueCategoryKeys_().map((key) => {
      const rule = DEFAULT_ISSUE_RULES[key];
      return [ISSUE_CATEGORIES[key].code, true, rule.value1, rule.value2, rule.description];
    }),
  format: (sheet, firstRow, rowCount) => {
    _settingsCheckboxes_(sheet, firstRow, rowCount, ISSUE_RULE_HEADERS.indexOf("Enabled") + 1);
    // plain text, so "0, 4" is never read as a number or a date
    sheet.getRange(firstRow, ISSUE_RULE_HEADERS.indexOf("Value 1") + 1, rowCount, 2).setNumberFormat("@");
  },
});

/**
 * rows of the Issue Rules table by rule code
 *
 * @returns {Map<string, {enabled: boolean, value1: unknown, value2: unknown}>} code → row
 */
function _readIssueRuleRows_() {
  return new Map(
    _readSection_(ISSUE_RULES_SECTION)
      .filter((row) => String(row[0]).trim() !== "")
      .map((row) => [String(row[0]).trim(), { enabled: row[1] === true, value1: row[2], value2: row[3] }]),
  );
}

/**
 * read every check's settings from the Issue Rules section of the Settings tab
 * values of disabled checks are not checked, a disabled check never adds issues
 *
 * @returns {IssueSettings} the settings
 * @throws {Error} if an enabled check has a value that cannot be read
 */
function _readIssueSettings_() {
  const rows = _readIssueRuleRows_();

  /** @param {IssueCategoryKey} key */
  const rowOf = (key) => rows.get(ISSUE_CATEGORIES[key].code) ?? { enabled: false, value1: "", value2: "" };

  /**
   * the value as text items, split on commas
   * @param {IssueCategoryKey} key
   * @param {1|2} which
   */
  const items = (key, which) => {
    const row = rowOf(key);
    const raw = which === 1 ? row.value1 : row.value2;
    const list = String(raw).split(",").map((item) => item.trim()).filter((item) => item !== "");
    if (row.enabled && list.length === 0) {
      throw new Error(`Issue Rules: ${ISSUE_CATEGORIES[key].code} needs Value ${which}`);
    }
    return list;
  };

  /**
   * the value as a list of numbers
   * @param {IssueCategoryKey} key
   * @param {1|2} which
   */
  const numbers = (key, which) =>
    items(key, which).map((item) => {
      const n = Number(item);
      if (rowOf(key).enabled && !Number.isFinite(n)) {
        throw new Error(
          `Issue Rules: ${ISSUE_CATEGORIES[key].code} Value ${which} must be a number, found "${item}"`,
        );
      }
      return n;
    });

  /**
   * the value as one number
   * @param {IssueCategoryKey} key
   * @param {1|2} which
   */
  const number = (key, which) => {
    const list = numbers(key, which);
    if (rowOf(key).enabled && list.length !== 1) {
      throw new Error(`Issue Rules: ${ISSUE_CATEGORIES[key].code} Value ${which} must be one number`);
    }
    return list.length === 0 ? NaN : list[0];
  };

  const keys = _issueCategoryKeys_();
  return {
    enabled: /** @type {Record<IssueCategoryKey, boolean>} */ (
      Object.fromEntries(keys.map((key) => [key, rowOf(key).enabled]))
    ),
    commentTotalAbove: number("totalWithoutComment", 1),
    commentTotalBelow: number("totalWithoutComment", 2),
    commentCategoryScores: numbers("categoryWithoutComment", 1),
    dangerousPlayKeywords: items("dangerousPlay", 1),
    lowScoreAtOrBelow: number("twoLowScores", 1),
    lowScoreCount: number("twoLowScores", 2),
    lowAverageBelow: number("lowAverage", 1),
    categoryMinimum: number("categoryMinimum", 1),
    singleLowScoreBelow: number("singleLowScore", 1),
    monitoringAverageBelow: number("monitoring", 1),
    monitoringBreaches: number("monitoring", 2),
  };
}
