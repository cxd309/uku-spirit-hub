/**
 * every data tab, in the order they appear
 * a function so every tab name constant has loaded before it is read
 *
 * @returns {string[]} tab names
 */
function _dataTabNames_() {
  return [ISSUES_SHEET, EVENTS_SHEET, RESPONSES_SHEET, TEAMS_SHEET, CLUBS_SHEET, CLUB_STATS_SHEET, CLUB_REPORT_SHEET];
}

/**
 * menu: set up the hub and open its settings
 * safe to run any number of times
 */
function showSettings() {
  _notify_(_withLock_(_showSettings_));
}

/**
 * create any missing tabs, refresh the Category choices,
 * then show the Settings tab on its own
 * tabs that already exist keep their contents
 * removes the empty default "Sheet1" once the hub tabs exist
 *
 * @returns {string} one-line summary
 */
function _showSettings_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const before = new Set(ss.getSheets().map((s) => s.getName()));

  _getIssuesSheet_();
  _getEventsSheet_();
  _getResponsesSheet_();
  _getTeamsSheet_();
  _getClubsSheet_();
  _getClubStatsSheet_();
  _getClubReportSheet_();
  const settings = _getSettingsSheet_();

  const blank = ss.getSheetByName("Sheet1");
  if (blank && blank.getLastRow() === 0 && blank.getLastColumn() === 0) ss.deleteSheet(blank);

  const note = _setCategoryChoices_();

  settings.showSheet();
  ss.setActiveSheet(settings);
  for (const name of _dataTabNames_()) ss.getSheetByName(name)?.hideSheet();

  const created = ss.getSheets().map((s) => s.getName()).filter((n) => !before.has(n));
  const summary = created.length === 0 ? "Settings" : `Created: ${created.join(", ")}`;
  return note === "" ? summary : `${summary}; ${note}`;
}

/**
 * show every data tab and hide the Settings tab
 * data tabs are shown first, so there is always a visible tab
 */
function _showDataTabs_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  for (const name of _dataTabNames_()) ss.getSheetByName(name)?.showSheet();
  ss.getSheetByName(SETTINGS_SHEET)?.hideSheet();
}
