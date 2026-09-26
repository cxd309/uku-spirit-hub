/**
 * settings in the Hub section of the Settings tab
 * keys are the names used in code
 */
const HUB_SETTINGS = Object.freeze({
  category: {
    label: "Category",
    default: "",
    description: "Folder inside the season folder that this hub covers, pick from the list",
  },
});

/**
 * the Hub section of the Settings tab
 */
const HUB_SECTION = Object.freeze<SettingsSection>({
  title: "Hub",
  description: "Which results this hub reads. Pick the Category once, changing it later would mix two categories",
  headers: SETTING_HEADERS,
  keyed: true,
  defaults: () => _settingRows_(HUB_SETTINGS),
  format: () => {},
});

/**
 * settings read from the Hub section
 */
interface HubSettings {
  /** category folder this hub reads, e.g. University */
  category: string;
}

/**
 * read the Hub section of the Settings tab
 *
 * @returns the settings
 * @throws {Error} if Category is blank
 */
function _readHubSettings_(): HubSettings {
  const category = _readSettingValues_(HUB_SECTION).get(HUB_SETTINGS.category.label) ?? "";
  if (category === "") {
    throw new Error(
      `Set "${HUB_SETTINGS.category.label}" on the ${SETTINGS_SHEET} tab (SpiritHub > SpiritHub Settings)`,
    );
  }
  return { category };
}

/**
 * make the Category value a dropdown of the folders in the season folder
 *
 * @param sheet  the Settings tab, from _getSettingsSheet_
 * @returns a note for the summary, "" when the dropdown was set
 */
function _setCategoryChoices_(sheet: GoogleAppsScript.Spreadsheet.Sheet): string {
  const names: string[] = [];
  try {
    const folders = _findSeasonFolder_().getFolders();
    while (folders.hasNext()) names.push(folders.next().getName());
  } catch (e) {
    return "Category list not updated, the hub is not inside a season folder";
  }
  if (names.length === 0) return "Category list not updated, no folders in the season folder";

  const columnA = _settingsColumnA_(sheet);
  const location = _findSection_(columnA, HUB_SECTION);
  if (!location || location.rowCount === 0) return "Category list not updated, Hub section not found";
  const labels = columnA
    .slice(location.firstRow - 1, location.firstRow - 1 + location.rowCount)
    .map((row) => String(row[0]).trim());
  const index = labels.indexOf(HUB_SETTINGS.category.label);
  if (index < 0) return "Category list not updated, Category row not found";

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(names.sort((a, b) => a.localeCompare(b)), true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(location.firstRow + index, 2).setDataValidation(rule);
  return "";
}

/**
 * document property holding the category the Tournaments tab was built from
 */
const SCANNED_CATEGORY_PROPERTY = "scannedCategory";

/**
 * stop if the Category has changed since the Tournaments tab was built
 * changing it would mark every tournament MISSING and mix two categories
 *
 * @param category         the Category setting now
 * @param tournamentCount  rows on the Tournaments tab
 * @throws {Error} if the category changed and there are tournaments
 */
function _checkCategory_(category: string, tournamentCount: number) {
  const scanned = PropertiesService.getDocumentProperties().getProperty(SCANNED_CATEGORY_PROPERTY);
  if (scanned && scanned !== category && tournamentCount > 0) {
    throw new Error(
      `This hub has ${scanned} tournaments but Category is now ${category}.`
        + `Tournament Category should not be changed after data has been imported`
        + `Set Category back to ${scanned}, or use a new hub for ${category}`,
    );
  }
}

/**
 * remember the category the Tournaments tab was built from
 *
 * @param category  the Category setting
 */
function _rememberCategory_(category: string) {
  PropertiesService.getDocumentProperties().setProperty(SCANNED_CATEGORY_PROPERTY, category);
}
