/**
 * header row of the Name Rules table
 * pattern is a regular expression removed from the end of team names
 * example shows the rule's effect
 * rules apply top to bottom when Enabled is ticked
 */
const NAME_RULE_HEADERS = Object.freeze(/** @type {const} */ (["Rule", "Pattern", "Example", "Enabled"]));

/**
 * rules written when the section is first created
 * only the seed and trailing-number rules start enabled
 * the rest are ticked per hub as needed
 * patterns are suffix-only (end in $) so letters inside a name are never removed
 */
const DEFAULT_NAME_RULES = Object.freeze([
  ["Seed suffix", "\\s*\\(\\d+\\)$", "KCL Men's 2 (SE2) (24) → KCL Men's 2 (SE 2)", true],
  ["Trailing brackets", "\\s*\\([^)]*\\)$", "KCL Men's 2 (SE2) → KCL Men's 2", true],
  ["Trailing number", "\\s+\\d+$", "KCL Men's 2 → KCL Men's", true],
  ["Women / Women's / W", "(?i)\\s+(women'?s?|w)$", "KCL Women's → KCL", false],
  ["Men / Men's / Open / M / O", "(?i)\\s+(men'?s?|open|m|o)$", "KCL Men's → KCL", false],
  ["Mixed / X", "(?i)\\s+(mixed|x)$", "KCL X → KCL", false],
]);

/**
 * the Name Rules section of the Settings tab
 * not keyed, so a default rule that is deleted stays deleted
 *
 * @type {SettingsSection}
 */
const NAME_RULES_SECTION = Object.freeze({
  title: "Name Rules",
  description: "Text patterns removed from the end of team names to suggest a club, applied top to bottom. "
    + "Tick Enabled to use a rule, or add rows for new rules, each with a name in the Rule column. "
    + "Teams and Clubs update straight away",
  headers: NAME_RULE_HEADERS,
  keyed: false,
  defaults: () => DEFAULT_NAME_RULES.map((rule) => [...rule]),
  format: (sheet, firstRow, rowCount) =>
    _settingsCheckboxes_(sheet, firstRow, rowCount, NAME_RULE_HEADERS.indexOf("Enabled") + 1),
});

/**
 * the enabled Name Rules as regular expressions, top to bottom
 * patterns use the same syntax as a Sheets REGEXREPLACE, a leading (?i) ignores case
 *
 * @returns {{regexes: RegExp[], invalid: string[]}} the rules, and names of rules whose pattern is not valid
 */
function _readNameRules_() {
  /** @type {RegExp[]} */
  const regexes = [];
  /** @type {string[]} */
  const invalid = [];
  for (const [rule, pattern, , enabled] of _readSection_(NAME_RULES_SECTION)) {
    const text = String(pattern).trim();
    if (enabled !== true || text === "") continue;
    const ignoreCase = text.startsWith("(?i)");
    try {
      regexes.push(new RegExp(ignoreCase ? text.slice(4) : text, ignoreCase ? "gi" : "g"));
    } catch (e) {
      invalid.push(String(rule));
    }
  }
  return { regexes, invalid };
}

/**
 * suggested club for a team: its name with each rule removed in turn
 * pure, the same as the old REGEXREPLACE formula
 * falls back to the team name if the rules would leave nothing
 *
 * @param {string}   team     team name
 * @param {RegExp[]} regexes  enabled rules, top to bottom
 * @returns {string} suggested club
 */
function _suggestClub_(team, regexes) {
  const club = regexes.reduce((name, regex) => name.replace(regex, "").replace(/ +/g, " ").trim(), team);
  return club === "" ? team : club;
}
