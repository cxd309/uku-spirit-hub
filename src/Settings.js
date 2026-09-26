/**
 * name of the Settings tab
 */
const SETTINGS_SHEET = "Settings";

/**
 * columns on the Settings tab, enough for the widest table
 */
const SETTINGS_WIDTH = 5;

/**
 * text for the info row at the top of the Settings tab
 */
const SETTINGS_INFO = `Settings for SpiritHub v${HUB_VERSION}\n\n`
  + "Each section has a title, a description and a table, only edit the tables and keep them free of blank rows\n"
  + "Name Rules changes apply straight away, everything else from the next refresh\n\n"
  + "Any refresh hides this tab, open it again with SpiritHub > SpiritHub Settings";

/**
 * header row of a table of single settings, like Hub and Spirit Award
 */
const SETTING_HEADERS = Object.freeze(/** @type {const} */ (["Setting", "Value", "Description"]));

/**
 * one setting in a table of single settings
 * `label` is the text in the Setting column, description explains it on the tab
 *
 * @typedef {Object} SettingDefinition
 * @property {string} label        text in the Setting column
 * @property {string} default      value written when the row is created
 * @property {string} description  shown in the Description column
 */

/**
 * one section of the Settings tab: a title, a description and a table
 *
 * @typedef {Object} SettingsSection
 * @property {string}            title        text in column A that marks the section, unique on the tab
 * @property {string}            description  shown under the title
 * @property {readonly string[]} headers      the table's header row
 * @property {boolean}           keyed        rows are named by their first column, missing ones are added back
 * @property {function(): unknown[][]} defaults  rows written when the section is created
 * @property {function(GoogleAppsScript.Spreadsheet.Sheet, number, number): void} format
 *   formats rows once written, given the tab, the first row and the row count
 */

/**
 * where a section's table is on the tab
 *
 * @typedef {Object} SectionLocation
 * @property {number} titleRow  row of the title, 1-based
 * @property {number} firstRow  first table row below the header, 1-based
 * @property {number} rowCount  how many table rows, may be 0
 */

/**
 * every section, in the order they appear
 * a function so every section constant has loaded before it is read
 *
 * @returns {SettingsSection[]} the sections
 */
function _settingsSections_() {
  return [HUB_SECTION, AWARD_SECTION, NAME_RULES_SECTION, ISSUE_RULES_SECTION];
}

/**
 * column A of the Settings tab, from row 1
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet  the Settings tab
 * @returns {unknown[][]} one single-value row per sheet row
 */
function _settingsColumnA_(sheet) {
  return sheet.getRange(1, 1, Math.max(sheet.getLastRow(), 1), 1).getValues();
}

/**
 * find a section by its title
 * pure, never modifies its arguments
 * the table starts below the title, description and header rows
 * and ends at the first blank row or the next section's title
 *
 * @param {unknown[][]}     columnA  column A of the Settings tab, from row 1
 * @param {SettingsSection} section  the section
 * @returns {SectionLocation|null} where the table is, null if the title is missing
 */
function _findSection_(columnA, section) {
  const titles = new Set(_settingsSections_().map((s) => s.title));
  /** @param {number} i */
  const text = (i) => String(columnA[i][0]).trim();
  const titleIndex = columnA.findIndex((_, i) => text(i) === section.title);
  if (titleIndex < 0) return null;

  const firstIndex = titleIndex + 3;
  let end = firstIndex;
  while (end < columnA.length && text(end) !== "" && !titles.has(text(end))) end++;
  return { titleRow: titleIndex + 1, firstRow: firstIndex + 1, rowCount: Math.max(end - firstIndex, 0) };
}

/**
 * write a section with its default rows below everything else on the tab
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet    the Settings tab
 * @param {SettingsSection}                    section  the section
 */
function _appendSection_(sheet, section) {
  const titleRow = Math.max(sheet.getLastRow(), INFO_ROW) + 2;
  const rows = section.defaults();
  const lastRow = titleRow + 2 + rows.length;
  if (sheet.getMaxRows() < lastRow) sheet.insertRowsAfter(sheet.getMaxRows(), lastRow - sheet.getMaxRows());

  sheet.getRange(titleRow, 1).setValue(section.title).setFontWeight("bold").setFontSize(12);
  sheet
    .getRange(titleRow + 1, 1, 1, SETTINGS_WIDTH)
    .merge()
    .setValue(section.description)
    .setWrap(true)
    .setVerticalAlignment("top");
  sheet.getRange(titleRow + 2, 1, 1, section.headers.length).setValues([[...section.headers]]).setFontWeight("bold");
  if (rows.length > 0) {
    sheet.getRange(titleRow + 3, 1, rows.length, section.headers.length).setValues(rows);
    section.format(sheet, titleRow + 3, rows.length);
  }
}

/**
 * make sure a section exists, and for keyed sections that every default row does
 * missing sections go at the bottom, missing rows at the end of their table
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet    the Settings tab
 * @param {SettingsSection}                    section  the section
 */
function _ensureSection_(sheet, section) {
  const location = _findSection_(_settingsColumnA_(sheet), section);
  if (!location) {
    _appendSection_(sheet, section);
    return;
  }
  if (!section.keyed) return;

  const present = new Set(
    location.rowCount === 0
      ? []
      : sheet.getRange(location.firstRow, 1, location.rowCount, 1).getValues().map((row) => String(row[0]).trim()),
  );
  const missing = section.defaults().filter((row) => !present.has(String(row[0])));
  if (missing.length === 0) return;

  const after = location.firstRow + location.rowCount - 1;
  sheet.insertRowsAfter(after, missing.length);
  sheet
    .getRange(after + 1, 1, missing.length, section.headers.length)
    .setValues(missing)
    .setFontWeight("normal");
  section.format(sheet, after + 1, missing.length);
}

/**
 * get the Settings tab, creating it and any missing section or keyed row on the way
 *
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} the Settings tab
 */
function _getSettingsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SETTINGS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(SETTINGS_SHEET);
    if (sheet.getMaxColumns() > SETTINGS_WIDTH) {
      sheet.deleteColumns(SETTINGS_WIDTH + 1, sheet.getMaxColumns() - SETTINGS_WIDTH);
    }
    sheet
      .getRange(INFO_ROW, 1, 1, SETTINGS_WIDTH)
      .merge()
      .setValue(SETTINGS_INFO)
      .setWrap(true)
      .setVerticalAlignment("top");
  }
  for (const section of _settingsSections_()) _ensureSection_(sheet, section);
  return sheet;
}

/**
 * the table rows of a section
 *
 * @param {SettingsSection} section  the section
 * @returns {unknown[][]} one array per row, as wide as the section's headers
 */
function _readSection_(section) {
  const sheet = _getSettingsSheet_();
  const location = _findSection_(_settingsColumnA_(sheet), section);
  if (!location || location.rowCount === 0) return [];
  return sheet.getRange(location.firstRow, 1, location.rowCount, section.headers.length).getValues();
}

/**
 * tick boxes in one column of newly written rows
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet     the Settings tab
 * @param {number}                             firstRow  first row
 * @param {number}                             rowCount  how many rows
 * @param {number}                             column    1-based column
 */
function _settingsCheckboxes_(sheet, firstRow, rowCount, column) {
  sheet
    .getRange(firstRow, column, rowCount, 1)
    .setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
}

/**
 * default rows for a table of single settings
 *
 * @param {Readonly<Record<string, SettingDefinition>>} settings  the settings, in display order
 * @returns {unknown[][]} one row per setting
 */
function _settingRows_(settings) {
  return Object.values(settings).map((s) => [s.label, s.default, s.description]);
}

/**
 * the values in a table of single settings
 *
 * @param {SettingsSection} section  the section
 * @returns {Map<string, string>} setting label → value, both trimmed
 */
function _readSettingValues_(section) {
  return new Map(_readSection_(section).map((row) => [String(row[0]).trim(), String(row[1]).trim()]));
}
