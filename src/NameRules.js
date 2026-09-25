/**
 * name of the Name Rules tab
 */
const NAME_RULES_SHEET = "Name Rules";

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
  ["Seed suffix", "\\s*\\(\\d+\\)$", "AUC(24) → AUC", true],
  ["Trailing number", "\\s+\\d+$", "Durham 2 → Durham", true],
  ["Women / Women's / W", "(?i)\\s+(women'?s?|w)$", "Bristol Women → Bristol", false],
  ["Men / Men's / Open / M / O", "(?i)\\s+(men'?s?|open|m|o)$", "SOLENT Open → SOLENT", false],
  ["Mixed / X", "(?i)\\s+(mixed|x)$", "LMU Mixed → LMU", false],
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
  const { sheet, created } = _getOrCreateSheet_(ss, NAME_RULES_SHEET, NAME_RULE_HEADERS);
  if (created) {
    const rows = DEFAULT_NAME_RULES.map((rule) => [...rule]);
    _fitSheet_(sheet, 1 + rows.length, NAME_RULE_HEADERS.length);
    sheet
      .getRange(2, 1, rows.length, NAME_RULE_HEADERS.length)
      .setValues(rows);
    sheet
      .getRange(2, NAME_RULE_HEADERS.indexOf("Enabled") + 1, rows.length, 1)
      .setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
  }
  return sheet;
}

/**
 * whole-column reference to a Name Rules column, below the header
 * e.g. 'Name Rules'!$B$2:$B
 *
 * @param {typeof NAME_RULE_HEADERS[number]} header  column header
 * @returns {string} the reference, for use inside formulas
 */
function _nameRuleColumn_(header) {
  const letter = _columnLetter_(NAME_RULE_HEADERS.indexOf(header) + 1);
  return `'${NAME_RULES_SHEET}'!$${letter}$2:$${letter}`;
}
