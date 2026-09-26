/**
 * add the SpiritHub menu when the spreadsheet is opened
 *
 * `onOpen` is a reserved name: Apps Script runs it on open
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("SpiritHub")
    .addItem("Refresh Tournaments", "refreshTournaments")
    .addItem("Refresh Results", "refreshResults")
    .addItem("Refresh Issues", "refreshIssues")
    .addItem("Refresh Club Statistics", "refreshClubStatistics")
    .addSeparator()
    .addItem("SpiritHub Settings", "showSettings")
    .addToUi();
}

/**
 * keep Teams and Clubs up to date when a person changes something that affects club names
 * `onEdit` is a reserved name, apps script runs it after every edit by a person
 *
 * Club Override on Teams: rebuild Clubs
 * Name Rules on Settings: work out every Suggested Club again, then rebuild Clubs
 * anything else is ignored
 * skips quietly if another run holds the lock
 *   Refresh Results rebuilds Teams and Clubs anyway
 *
 * @param e  the edit event
 */
function onEdit(e: GoogleAppsScript.Events.SheetsOnEdit) {
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();

  const overrideColumn = TEAM_KEYS.indexOf("clubOverride") + 1;
  const touchesOverride = sheetName === TEAMS_SHEET
    && e.range.getColumn() <= overrideColumn
    && e.range.getLastColumn() >= overrideColumn;

  const nameRules = sheetName === SETTINGS_SHEET ? _findSection_(_settingsColumnA_(sheet), NAME_RULES_SECTION) : null;
  const touchesNameRules = nameRules !== null
    && e.range.getRow() <= nameRules.firstRow + nameRules.rowCount - 1
    && e.range.getLastRow() >= nameRules.firstRow;

  if (!touchesOverride && !touchesNameRules) return;

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(LOCK_WAIT_MS)) return;
  try {
    if (touchesNameRules) _refreshSuggestedClubs_();
    _rebuildClubsFromSheet_();
  } finally {
    lock.releaseLock();
  }
}

/**
 * how long to wait for another run to finish before giving up
 */
const LOCK_WAIT_MS = 1000;

/**
 * run a function while holding this spreadsheet's script lock
 * so two runs (by same or different people) can never conflict
 * @param work the function to run
 * @returns whatever `work` returns
 * @throws {Error} if another run still holds the lock after LOCK_WAIT_MS
 */
function _withLock_<T>(work: () => T): T {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(LOCK_WAIT_MS)) {
    throw new Error("Another SpiritHub run is in progress. Please try again in a minute");
  }
  try {
    return work();
  } finally {
    lock.releaseLock();
  }
}

/**
 * show a short message in the bottom-right corner of the spreadsheet
 * and log it
 *
 * @param message text to show
 */
function _notify_(message: string) {
  console.log(message);
  SpreadsheetApp.getActiveSpreadsheet().toast(message, "SpiritHub", 10);
}

/**
 * run a function and log how long it took, to find slow steps
 * shows in the Apps Script Executions log
 *
 * @param label  name of the step
 * @param work   the step
 * @returns whatever `work` returns
 */
function _timed_<T>(label: string, work: () => T): T {
  const start = Date.now();
  try {
    return work();
  } finally {
    console.log(`${label} ${((Date.now() - start) / 1000).toFixed(1)}s`);
  }
}

/**
 * run a refresh from the menu
 * shows every data tab and hides Settings first, then runs the work under the lock
 *
 * @param work  the refresh, returning its summary
 */
function _runRefresh_(work: () => string) {
  _notify_(_withLock_(() => {
    _showDataTabs_();
    return work();
  }));
}

/**
 * menu: scan the category folder and update the Tournaments tab
 */
function refreshTournaments() {
  _runRefresh_(_refreshTournaments_);
}

/**
 * menu: import every tournament marked NEW or REFRESH
 */
function refreshResults() {
  _runRefresh_(_refreshResults_);
}

/**
 * menu: add any new issues from what is already in the spreadsheet
 */
function refreshIssues() {
  _runRefresh_(() => _timed_("issues", _refreshIssues_));
}

/**
 * menu: recalculate the Club Statistics tab from what is already in the spreadsheet
 * only ever run from the menu
 */
function refreshClubStatistics() {
  _runRefresh_(_refreshClubStatistics_);
}
