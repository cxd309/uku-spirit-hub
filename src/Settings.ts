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
const SETTING_HEADERS = Object.freeze(["Setting", "Value", "Description"] as const);

/**
 * one setting in a table of single settings
 * `label` is the text in the Setting column, description explains it on the tab
 */
interface SettingDefinition {
  /** text in the Setting column */
  label: string;
  /** value written when the row is created */
  default: string;
  /** shown in the Description column */
  description: string;
}

/**
 * one section of the Settings tab: a title, a description and a table
 */
interface SettingsSection {
  /** text in column A that marks the section, unique on the tab */
  title: string;
  /** shown under the title */
  description: string;
  /** the table's header row */
  headers: readonly string[];
  /** rows are named by their first column, missing ones are added back */
  keyed: boolean;
  /** rows written when the section is created */
  defaults: () => unknown[][];
  /**
   * formats rows once written, given the tab, the first row and the row count
   */
  format: (sheet: GoogleAppsScript.Spreadsheet.Sheet, firstRow: number, rowCount: number) => void;
}

/**
 * where a section's table is on the tab
 */
interface SectionLocation {
  /** row of the title, 1-based */
  titleRow: number;
  /** first table row below the header, 1-based */
  firstRow: number;
  /** how many table rows, may be 0 */
  rowCount: number;
}

/**
 * every section, in the order they appear
 * a function so every section constant has loaded before it is read
 *
 * @returns the sections
 */
function _settingsSections_(): SettingsSection[] {
  return [HUB_SECTION, AWARD_SECTION, NAME_RULES_SECTION, ISSUE_RULES_SECTION];
}

/**
 * column A of the Settings tab, from row 1
 *
 * @param sheet  the Settings tab
 * @returns one single-value row per sheet row
 */
function _settingsColumnA_(sheet: GoogleAppsScript.Spreadsheet.Sheet): unknown[][] {
  return sheet.getRange(1, 1, Math.max(sheet.getLastRow(), 1), 1).getValues();
}

/**
 * find a section by its title
 * pure, never modifies its arguments
 * the table starts below the title, description and header rows
 * and ends at the first blank row or the next section's title
 *
 * @param columnA  column A of the Settings tab, from row 1
 * @param section  the section
 * @returns where the table is, null if the title is missing
 */
function _findSection_(columnA: unknown[][], section: SettingsSection): SectionLocation | null {
  const titles = new Set(_settingsSections_().map((s) => s.title));
  const text = (i: number) => String(columnA[i][0]).trim();
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
 * @param sheet    the Settings tab
 * @param section  the section
 */
function _appendSection_(sheet: GoogleAppsScript.Spreadsheet.Sheet, section: SettingsSection) {
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
 * @param sheet    the Settings tab
 * @param section  the section
 * @param columnA  column A of the Settings tab as it is now, from row 1
 * @returns true if anything was written, so column A must be read again
 */
function _ensureSection_(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  section: SettingsSection,
  columnA: unknown[][],
): boolean {
  const location = _findSection_(columnA, section);
  if (!location) {
    _appendSection_(sheet, section);
    return true;
  }
  if (!section.keyed) return false;

  const present = new Set(
    columnA.slice(location.firstRow - 1, location.firstRow - 1 + location.rowCount).map((row) => String(row[0]).trim()),
  );
  const missing = section.defaults().filter((row) => !present.has(String(row[0])));
  if (missing.length === 0) return false;

  const after = location.firstRow + location.rowCount - 1;
  sheet.insertRowsAfter(after, missing.length);
  sheet
    .getRange(after + 1, 1, missing.length, section.headers.length)
    .setValues(missing)
    .setFontWeight("normal");
  section.format(sheet, after + 1, missing.length);
  return true;
}

/**
 * get the Settings tab, creating it and any missing section or keyed row on the way
 *
 * @returns the Settings tab
 */
function _getSettingsSheet_(): GoogleAppsScript.Spreadsheet.Sheet {
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
  // column A is read once, and again only after a section or row has been added
  let columnA = _settingsColumnA_(sheet);
  for (const section of _settingsSections_()) {
    if (_ensureSection_(sheet, section, columnA)) columnA = _settingsColumnA_(sheet);
  }
  return sheet;
}

/**
 * every value on the Settings tab, in one read
 * the tab is created with every section if it does not exist yet
 *
 * @returns one array per sheet row from row 1, all the same width
 */
function _readSettingsValues_(): unknown[][] {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SETTINGS_SHEET) ?? _getSettingsSheet_();
  return sheet.getDataRange().getValues();
}

/**
 * the table rows of a section, from values already read
 * pure, never modifies its arguments
 *
 * @param values   every value on the Settings tab, from _readSettingsValues_
 * @param section  the section
 * @returns one array per row, as wide as the section's headers, null if the section is missing
 */
function _sectionRows_(values: unknown[][], section: SettingsSection): unknown[][] | null {
  const location = _findSection_(values, section);
  if (!location) return null;
  return values
    .slice(location.firstRow - 1, location.firstRow - 1 + location.rowCount)
    .map((row) => Array.from({ length: section.headers.length }, (_, i) => row[i] ?? ""));
}

/**
 * whether a section is complete: present, and for keyed sections with every default row
 * pure, never modifies its arguments
 *
 * @param rows     the section's rows from _sectionRows_
 * @param section  the section
 * @returns true if nothing needs adding
 */
function _isSectionComplete_(rows: unknown[][] | null, section: SettingsSection): boolean {
  if (!rows) return false;
  if (!section.keyed) return true;
  const present = new Set(rows.map((row) => String(row[0]).trim()));
  return section.defaults().every((row) => present.has(String(row[0])));
}

/**
 * the table rows of a section
 * one read of the Settings tab, a missing section or keyed row is added back first (then read again)
 * so checks and settings added in later versions still appear on their own
 *
 * @param section  the section
 * @returns one array per row, as wide as the section's headers
 */
function _readSection_(section: SettingsSection): unknown[][] {
  const rows = _sectionRows_(_readSettingsValues_(), section);
  if (_isSectionComplete_(rows, section)) return rows ?? [];
  _getSettingsSheet_();
  return _sectionRows_(_readSettingsValues_(), section) ?? [];
}

/**
 * tick boxes in one column of newly written rows
 *
 * @param sheet     the Settings tab
 * @param firstRow  first row
 * @param rowCount  how many rows
 * @param column    1-based column
 */
function _settingsCheckboxes_(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  firstRow: number,
  rowCount: number,
  column: number,
) {
  sheet
    .getRange(firstRow, column, rowCount, 1)
    .setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
}

/**
 * default rows for a table of single settings
 *
 * @param settings  the settings, in display order
 * @returns one row per setting
 */
function _settingRows_(settings: Readonly<Record<string, SettingDefinition>>): unknown[][] {
  return Object.values(settings).map((s) => [s.label, s.default, s.description]);
}

/**
 * the values in a table of single settings
 *
 * @param section  the section
 * @returns setting label → value, both trimmed
 */
function _readSettingValues_(section: SettingsSection): Map<string, string> {
  return new Map(_readSection_(section).map((row) => [String(row[0]).trim(), String(row[1]).trim()]));
}
