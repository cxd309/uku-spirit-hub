/**
 * settings in the Spirit Award section of the Settings tab
 * keys are the names used in code
 */
const AWARD_SETTINGS = Object.freeze({
  minimumTournaments: {
    label: "Award Minimum Tournaments",
    default: "4",
    description: "Minimum number of tournaments a club must enter to qualify for the spirit award",
  },
  includeInternational: {
    label: "Include International",
    default: "TRUE",
    description: "Count tournaments ticked International on the Tournaments tab towards the spirit award",
  },
});

/**
 * the Spirit Award section of the Settings tab
 */
const AWARD_SECTION = Object.freeze<SettingsSection>({
  title: "Spirit Award",
  description: "How clubs qualify for the spirit award on the Club Statistics tab",
  headers: SETTING_HEADERS,
  keyed: true,
  editable: () => [SETTING_HEADERS.indexOf("Value")],
  defaults: () => _settingRows_(AWARD_SETTINGS),
  format: (sheet, firstRow, rowCount) => {
    const labels = sheet.getRange(firstRow, 1, rowCount, 1).getValues().map((row) => String(row[0]).trim());
    labels.forEach((label, i) => {
      if (label === AWARD_SETTINGS.includeInternational.label) _settingsCheckboxes_(sheet, firstRow + i, 1, 2);
    });
  },
});

/**
 * settings read from the Spirit Award section
 */
interface AwardSettings {
  /** tournaments a club must enter to qualify */
  minimumTournaments: number;
  /** whether tournaments ticked International count towards the award */
  includeInternational: boolean;
}

/**
 * read the Spirit Award section of the Settings tab
 *
 * @returns the settings
 * @throws {Error} if a value is not a whole number
 */
function _readAwardSettings_(): AwardSettings {
  const values = _readSettingValues_(AWARD_SECTION);
  const label = AWARD_SETTINGS.minimumTournaments.label;
  const raw = values.get(label) ?? "";
  const minimumTournaments = Number(raw);
  if (raw === "" || !Number.isInteger(minimumTournaments) || minimumTournaments < 0) {
    throw new Error(`"${label}" on the ${SETTINGS_SHEET} tab must be a whole number, found "${raw}"`);
  }
  const includeInternational = (values.get(AWARD_SETTINGS.includeInternational.label) ?? "").toUpperCase() === "TRUE";
  return { minimumTournaments, includeInternational };
}

/**
 * whether a tournament's scores count towards the spirit award
 * pure, no google calls
 *
 * @param event     the tournament
 * @param settings  the Spirit Award settings
 * @returns true if Include Results is ticked, and it is not International or international tournaments count
 */
function _eventCountsForAward_(event: EventRecord, settings: AwardSettings): boolean {
  return event.includeResults && (!event.international || settings.includeInternational);
}
