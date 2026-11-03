/**
 * settings in the Hub section of the Settings tab
 * keys are the names used in code
 */
const HUB_SETTINGS = Object.freeze({
  scheduleManagerUrl: {
    label: "Schedule Manager URL",
    default: "",
    description: "Link to the Schedule Manager spreadsheet, tournaments are read from its Schedule List tab",
  },
  startDate: {
    label: "Start Date",
    default: "",
    description: "Only tournaments starting on or after this date are listed",
  },
  endDate: {
    label: "End Date",
    default: "",
    description: "Only tournaments starting on or before this date are listed",
  },
});

/**
 * labels of the Hub settings that hold a date
 */
const HUB_DATE_LABELS = Object.freeze([HUB_SETTINGS.startDate.label, HUB_SETTINGS.endDate.label]);

/**
 * the Hub section of the Settings tab
 */
const HUB_SECTION = Object.freeze<SettingsSection>({
  title: "Hub",
  description: "Where this hub reads its tournaments from, and which dates it covers",
  headers: SETTING_HEADERS,
  keyed: true,
  defaults: () => _settingRows_(HUB_SETTINGS),
  format: (sheet, firstRow, rowCount) => {
    const labels = sheet.getRange(firstRow, 1, rowCount, 1).getValues().map((row) => String(row[0]).trim());
    const rule = SpreadsheetApp.newDataValidation().requireDate().build();
    labels.forEach((label, i) => {
      if (HUB_DATE_LABELS.includes(label)) {
        sheet.getRange(firstRow + i, 2).setDataValidation(rule).setNumberFormat("yyyy-mm-dd");
      }
    });
  },
});

/**
 * settings read from the Hub section
 */
interface HubSettings {
  /** link to the Schedule Manager spreadsheet */
  scheduleManagerUrl: string;
  /** first start date listed */
  startDate: Date;
  /** last start date listed */
  endDate: Date;
}

/**
 * read the Hub section of the Settings tab
 * the values are read as they are, not as text, so the dates stay dates
 *
 * @returns the settings
 * @throws {Error} if a setting is blank or not a date, or End Date is before Start Date
 */
function _readHubSettings_(): HubSettings {
  const values = new Map(_readSection_(HUB_SECTION).map((row) => [String(row[0]).trim(), row[1]]));
  const where = `on the ${SETTINGS_SHEET} tab (SpiritHub > SpiritHub Settings)`;

  const scheduleManagerUrl = String(values.get(HUB_SETTINGS.scheduleManagerUrl.label) ?? "").trim();
  if (scheduleManagerUrl === "") throw new Error(`Set "${HUB_SETTINGS.scheduleManagerUrl.label}" ${where}`);

  const date = (label: string): Date => {
    const value = values.get(label);
    if (!(value instanceof Date)) throw new Error(`Set "${label}" to a date ${where}`);
    return value;
  };
  const startDate = date(HUB_SETTINGS.startDate.label);
  const endDate = date(HUB_SETTINGS.endDate.label);
  if (endDate.getTime() < startDate.getTime()) {
    throw new Error(`"${HUB_SETTINGS.endDate.label}" is before "${HUB_SETTINGS.startDate.label}" ${where}`);
  }
  return { scheduleManagerUrl, startDate, endDate };
}
