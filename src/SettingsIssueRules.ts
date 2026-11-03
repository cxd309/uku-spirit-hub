/**
 * header row of the Issue Rules table
 */
const ISSUE_RULE_HEADERS = Object.freeze(
  ["Rule", "Enabled", "Value 1", "Value 2", "Description"] as const,
);

/**
 * the row written for each check when it is missing from the section
 * value 1 and value 2 mean different things per check, the description says what
 * lists are written as comma separated text
 */
const DEFAULT_ISSUE_RULES: Readonly<
  Record<IssueCategoryKey, { value1: number | string; value2: number | string; description: string }>
> = Object.freeze({
  lowScore: {
    value1: 6,
    value2: 2,
    description: "Play, policy: received a total of Value 1 or below. "
      + "Details say when the team has Value 2 or more of them at the tournament, the policy trigger",
  },
  dangerousPlay: {
    value1: "dangerous, danger, reckless, unsafe",
    value2: "",
    description: "Play, policy: a comment received mentions dangerous play, any of the words in Value 1",
  },
  cheating: {
    value1: "cheat, cheating, intentional",
    value2: "",
    description: "Play: a comment received mentions cheating, any of the words in Value 1",
  },
  harassment: {
    value1: "abuse, abusive, swear, swearing, unwarranted, harass, harassment, insult, insulting",
    value2: "",
    description: "Play: a comment received mentions harassment or abuse, any of the words in Value 1",
  },
  categoryMinimum: {
    value1: 0,
    value2: "",
    description: "Play: received Value 1 in any category",
  },
  totalWithoutComment: {
    value1: 14,
    value2: 6,
    description: "Scoring, policy: gave a total above Value 1 or below Value 2 without a comment",
  },
  categoryWithoutComment: {
    value1: "0, 4",
    value2: "",
    description: "Scoring, policy: gave any of Value 1 in a category without a comment",
  },
  deletedRow: {
    value1: "",
    value2: "",
    description: "Scoring: a form is on the Deleted Forms tab, with the chair's reason",
  },
  selfScore: {
    value1: "",
    value2: "",
    description: "Scoring: a team scored itself",
  },
  noSuchGame: {
    value1: "",
    value2: "",
    description: "Scoring: a form for two teams with no game between them on the schedule",
  },
  duplicate: {
    value1: "",
    value2: "",
    description: "Scoring: more forms to an opponent than games played, the one with a comment then the latest counts",
  },
  noScores: {
    value1: "",
    value2: "",
    description: "Tournament, policy: a team submitted no scores at a tournament",
  },
  underSubmitted: {
    value1: 50,
    value2: "",
    description: "Tournament, policy: a team submitted scores for fewer than Value 1 % of the games it played",
  },
  immediateAverage: {
    value1: 8,
    value2: "",
    description: "Tournament, policy: a team averaged below Value 1 at a tournament",
  },
  monitoringAverage: {
    value1: 9,
    value2: "",
    description: "Tournament, policy: a team averaged below Value 1 at a tournament, but not below IMMEDIATE-AVERAGE",
  },
});

/**
 * thresholds and switches for every issue check, read from the Issue Rules settings
 */
interface IssueSettings {
  /** which checks add issues */
  enabled: Record<IssueCategoryKey, boolean>;
  /** a total at or below this is a low score */
  lowScoreAtOrBelow: number;
  /** this many low scores at a tournament meets the policy */
  lowScoreCount: number;
  /** words that flag a comment, for each comment keyword check */
  commentKeywords: Record<typeof COMMENT_KEYWORD_CHECKS[number], string[]>;
  /** receiving this in a category is an issue */
  categoryMinimum: number;
  /** total above this needs a comment */
  commentTotalAbove: number;
  /** total below this needs a comment */
  commentTotalBelow: number;
  /** category scores that need a comment */
  commentCategoryScores: number[];
  /** submitting for fewer than this percentage of played games is an issue */
  submittedPercentBelow: number;
  /** an average below this at a tournament needs immediate action */
  immediateAverageBelow: number;
  /** an average below this at a tournament counts towards monitoring */
  monitoringAverageBelow: number;
}

/**
 * every check key, in the order of ISSUE_CATEGORIES
 *
 * @returns the keys
 */
function _issueCategoryKeys_(): IssueCategoryKey[] {
  return Object.keys(ISSUE_CATEGORIES) as IssueCategoryKey[];
}

/**
 * the cells of an Issue Rules row that are filled in: Enabled, and each value the rule uses
 * a rule uses a value when its default is not blank, rules this version does not know only have Enabled
 *
 * @param code  the Rule column, e.g. LOW-SCORE
 * @returns 0-based columns
 */
function _issueRuleEditable_(code: string): number[] {
  const column = (header: typeof ISSUE_RULE_HEADERS[number]) => ISSUE_RULE_HEADERS.indexOf(header);
  const key = _issueCategoryKeys_().find((k) => ISSUE_CATEGORIES[k].code === code);
  if (!key) return [column("Enabled")];
  const rule = DEFAULT_ISSUE_RULES[key];
  return [
    column("Enabled"),
    ...(rule.value1 === "" ? [] : [column("Value 1")]),
    ...(rule.value2 === "" ? [] : [column("Value 2")]),
  ];
}

/**
 * the Issue Rules section of the Settings tab
 * keyed by rule code, so a missing rule is added back with its defaults
 * and checks added in later versions appear on their own
 */
const ISSUE_RULES_SECTION = Object.freeze<SettingsSection>({
  title: "Issue Rules",
  description: "Settings for each issue check, used by Refresh Issues. "
    + "Untick Enabled to stop adding new issues of that type. "
    + "Value 1 and Value 2 are thresholds, the description says what each means, lists are comma separated. "
    + "Do not change the Rule column. Rules no longer used by this version can be deleted",
  headers: ISSUE_RULE_HEADERS,
  keyed: true,
  editable: _issueRuleEditable_,
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
 * @returns code → row
 */
function _readIssueRuleRows_(): Map<string, { enabled: boolean; value1: unknown; value2: unknown }> {
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
 * @returns the settings
 * @throws {Error} if an enabled check has a value that cannot be read
 */
function _readIssueSettings_(): IssueSettings {
  const rows = _readIssueRuleRows_();

  const rowOf = (key: IssueCategoryKey) =>
    rows.get(ISSUE_CATEGORIES[key].code) ?? { enabled: false, value1: "", value2: "" };

  /**
   * the value as text items, split on commas
   */
  const items = (key: IssueCategoryKey, which: 1 | 2) => {
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
   */
  const numbers = (key: IssueCategoryKey, which: 1 | 2) =>
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
   */
  const number = (key: IssueCategoryKey, which: 1 | 2) => {
    const list = numbers(key, which);
    if (rowOf(key).enabled && list.length !== 1) {
      throw new Error(`Issue Rules: ${ISSUE_CATEGORIES[key].code} Value ${which} must be one number`);
    }
    return list.length === 0 ? NaN : list[0];
  };

  const keys = _issueCategoryKeys_();
  return {
    enabled: Object.fromEntries(keys.map((key) => [key, rowOf(key).enabled])) as Record<IssueCategoryKey, boolean>,
    lowScoreAtOrBelow: number("lowScore", 1),
    lowScoreCount: number("lowScore", 2),
    commentKeywords: {
      dangerousPlay: items("dangerousPlay", 1),
      cheating: items("cheating", 1),
      harassment: items("harassment", 1),
    },
    categoryMinimum: number("categoryMinimum", 1),
    commentTotalAbove: number("totalWithoutComment", 1),
    commentTotalBelow: number("totalWithoutComment", 2),
    commentCategoryScores: numbers("categoryWithoutComment", 1),
    submittedPercentBelow: number("underSubmitted", 1),
    immediateAverageBelow: number("immediateAverage", 1),
    monitoringAverageBelow: number("monitoringAverage", 1),
  };
}
