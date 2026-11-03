/**
 * every data tab, in the order they appear
 * a function so every tab name constant has loaded before it is read
 *
 * @returns tab names
 */
function _dataTabNames_(): string[] {
  return [
    ISSUES_SHEET,
    EVENTS_SHEET,
    GAMES_SHEET,
    RESPONSES_SHEET,
    DELETED_FORMS_SHEET,
    TEAMS_SHEET,
    CLUBS_SHEET,
    CLUB_CONTACTS_SHEET,
    CLUB_STATS_SHEET,
    CLUB_REPORT_SHEET,
  ];
}

/**
 * tabs shown with the Settings tab and hidden by every refresh
 * a function so every tab name constant has loaded before it is read
 *
 * @returns tab names
 */
function _settingsTabNames_(): string[] {
  return [SETTINGS_SHEET, EMAIL_TEMPLATES_SHEET];
}

/**
 * menu: set up the hub and open its settings
 * safe to run any number of times
 */
function showSettings() {
  _notify_(_withLock_(_showSettings_));
}

/**
 * create any missing tabs,
 * then show the Settings and Email Templates tabs on their own
 * tabs that already exist keep their contents
 * removes the empty default "Sheet1" once the hub tabs exist
 *
 * @returns one-line summary
 */
function _showSettings_(): string {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const before = new Set(ss.getSheets().map((s) => s.getName()));

  _getIssuesSheet_();
  _getEventsSheet_();
  _getGamesSheet_();
  _getResponsesSheet_();
  _getDeletedFormsSheet_();
  _getTeamsSheet_();
  _getClubsSheet_();
  _getClubContactsSheet_();
  _getClubStatsSheet_();
  _getClubReportSheet_();
  const settings = _getSettingsSheet_();
  _getEmailTemplatesSheet_();

  const blank = ss.getSheetByName("Sheet1");
  if (blank && blank.getLastRow() === 0 && blank.getLastColumn() === 0) ss.deleteSheet(blank);

  for (const name of _settingsTabNames_()) ss.getSheetByName(name)?.showSheet();
  ss.setActiveSheet(settings);
  for (const name of _dataTabNames_()) ss.getSheetByName(name)?.hideSheet();

  const created = ss.getSheets().map((s) => s.getName()).filter((n) => !before.has(n));
  const summary = created.length === 0 ? "Settings" : `Created: ${created.join(", ")}`;
  return summary;
}

/**
 * show every data tab and hide the Settings and Email Templates tabs
 * data tabs are shown first, so there is always a visible tab
 */
function _showDataTabs_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  for (const name of _dataTabNames_()) ss.getSheetByName(name)?.showSheet();
  for (const name of _settingsTabNames_()) ss.getSheetByName(name)?.hideSheet();
}
