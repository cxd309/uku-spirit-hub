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
});

/**
 * the Spirit Award section of the Settings tab
 */
const AWARD_SECTION = Object.freeze<SettingsSection>({
  title: "Spirit Award",
  description: "How clubs qualify for the spirit award on the Club Statistics tab",
  headers: SETTING_HEADERS,
  keyed: true,
  defaults: () => _settingRows_(AWARD_SETTINGS),
  format: () => {},
});

/**
 * settings read from the Spirit Award section
 */
interface AwardSettings {
  /** tournaments a club must enter to qualify */
  minimumTournaments: number;
}

/**
 * read the Spirit Award section of the Settings tab
 *
 * @returns the settings
 * @throws {Error} if a value is not a whole number
 */
function _readAwardSettings_(): AwardSettings {
  const label = AWARD_SETTINGS.minimumTournaments.label;
  const raw = _readSettingValues_(AWARD_SECTION).get(label) ?? "";
  const minimumTournaments = Number(raw);
  if (raw === "" || !Number.isInteger(minimumTournaments) || minimumTournaments < 0) {
    throw new Error(`"${label}" on the ${SETTINGS_SHEET} tab must be a whole number, found "${raw}"`);
  }
  return { minimumTournaments };
}
