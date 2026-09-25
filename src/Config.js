/**
 * name of the Config tab
 */

const CONFIG_SHEET = "Config";

/**
 * text for the info row above the Config table
 * what the tab is, what to edit, how it refreshes
 */
const CONFIG_INFO = `Settings for SpiritHub v${HUB_VERSION}\n\n`
  + "Only edit the values column";

/**
 * header row of the Config tab
 */
const CONFIG_HEADERS = Object.freeze([
  "Setting",
  "Value",
  "Description",
]);

/**
 * every setting the Hub reads
 * keys are the names used in code
 * `label` is the text in the setting column
 * description explains the setting on the tab
 */
const CONFIG_SETTINGS = Object.freeze(
  /** @type {const} */ ({
    category: {
      label: "Category",
      default: "",
      description: "Folder inside the season folder that this Hub covers, e.g. University, Club, ...",
    },
    awardMinimumTournaments: {
      label: "Award Minimum Tournaments",
      default: "3",
      description: "Minimum number of tournaments a club must enter to qualify for the spirit award",
    },
  }),
);

/**
 * settings read from the Config tab
 *
 * @typedef {Record<keyof typeof CONFIG_SETTINGS, string>} HubConfig
 */

/**
 * get the config tab
 * create if needed and add row for any setting not filled
 *
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} config tab
 */
function _getConfigSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const { sheet } = _getOrCreateSheet_(ss, CONFIG_SHEET, CONFIG_HEADERS, CONFIG_INFO);
  const present = new Set(_readConfigRows_(sheet).map((row) => row.label));
  const missing = Object.values(CONFIG_SETTINGS).filter((s) => !present.has(s.label));
  if (missing.length > 0) {
    const firstEmpty = sheet.getLastRow() + 1;
    _fitSheet_(sheet, firstEmpty + missing.length - 1, CONFIG_HEADERS.length);
    sheet
      .getRange(firstEmpty, 1, missing.length, CONFIG_HEADERS.length)
      .setValues(missing.map((s) => [s.label, s.default, s.description]));
  }
  return sheet;
}

/**
 * Read the Setting and Value columns of the Config tab
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet the config tab
 * @returns {{label:string, value:string}[]} one entry per non-empty row
 */
function _readConfigRows_(sheet) {
  const rowCount = sheet.getLastRow() - HEADER_ROW;
  if (rowCount < 1) return [];
  return sheet
    .getRange(DATA_ROW, 1, rowCount, 2)
    .getValues()
    .map(([label, value]) => ({ label: String(label).trim(), value: String(value).trim() }))
    .filter((row) => row.label !== "");
}

/**
 * read the Hub's settings from the Config tab
 * @returns {HubConfig} the settings
 * @throws {Error} if a required setting is blank
 */
function _readConfig_() {
  const values = new Map(_readConfigRows_(_getConfigSheet_()).map((row) => [row.label, row.value]));
  const category = values.get(CONFIG_SETTINGS.category.label) ?? "";
  if (category === "") {
    throw new Error(`Set "${CONFIG_SETTINGS.category.label}" on the ${CONFIG_SHEET} tab`);
  }
  return {
    category: category,
    awardMinimumTournaments: values.get(CONFIG_SETTINGS.awardMinimumTournaments.label) ?? "",
  };
}
