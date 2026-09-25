/**
 * name of the Name Rules tab
 */
const NAME_RULES_SHEET = "Name Rules";

/**
 * text for the info row above the NameRules table
 * what the tab is, what to edit, how it refreshes
 */
const NAME_RULES_INFO = "Text patterns removed from the end of team names to suggest a club, applied top to bottom.\n\n"
  + "You can tick enabled to enable a rule or add extra rows with additional rules.\n\n"
  + "Teams and Clubs tab will refresh instantly when this is changed.";

/**
 * header row of the Name Rules tab
 * pattern is a regular expression removed from the end of team names
 * example shows the rule's effect
 * rules apply top to bottom when Enabled is ticked
 */
const NAME_RULE_HEADERS = Object.freeze(/** @type {const} */ (["Rule", "Pattern", "Example", "Enabled"]));

/**
 * rules written when the tab is first created
 * only the seed and trailing-number rules start enabled
 * the rest are ticked per hub as needed
 * patterns are suffix-only (end in $) so letters inside a name are never removed
 */
const DEFAULT_NAME_RULES = Object.freeze([
  ["Seed suffix", "\\s*\\(\\d+\\)$", "KCL Men's 2 (SE2) (24) → KCL Men's 2 (SE 2)", true],
  ["Trailing brackets", "\\s*\\([^)]*\\)$", "KCL Men's 2 (SE2) → KCL Men's 2", true],
  ["Trailing number", "\\s+\\d+$", "KCL Men's 2 → KCL", true],
  ["Women / Women's / W", "(?i)\\s+(women'?s?|w)$", "KCL Women's → KCL", false],
  ["Men / Men's / Open / M / O", "(?i)\\s+(men'?s?|open|m|o)$", "KCL Men's → KCL", false],
  ["Mixed / X", "(?i)\\s+(mixed|x)$", "KCL X → KCL", false],
]);

/**
 * get the Name Rules tab
 * creating it with the default rules on first use
 * after creation the script never changes it
 *
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}  the Name Rules tab
 */
function _getNameRulesSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const { sheet, created } = _getOrCreateSheet_(ss, NAME_RULES_SHEET, NAME_RULE_HEADERS, NAME_RULES_INFO);
  if (created) {
    const rows = DEFAULT_NAME_RULES.map((rule) => [...rule]);
    _fitSheet_(sheet, HEADER_ROW + rows.length, NAME_RULE_HEADERS.length);
    sheet
      .getRange(DATA_ROW, 1, rows.length, NAME_RULE_HEADERS.length)
      .setValues(rows);
    sheet
      .getRange(DATA_ROW, NAME_RULE_HEADERS.indexOf("Enabled") + 1, rows.length, 1)
      .setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
  }
  return sheet;
}

/**
 * whole-column reference to a Name Rules column, below the header
 * e.g. 'Name Rules'!$B$3:$B
 *
 * @param {typeof NAME_RULE_HEADERS[number]} header  column header
 * @returns {string} the reference, for use inside formulas
 */
function _nameRuleColumn_(header) {
  return _columnBelowHeader_(NAME_RULES_SHEET, NAME_RULE_HEADERS.indexOf(header) + 1);
}
