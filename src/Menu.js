/**
 * add the SpiritHub menu when the spreadsheet is opened
 *
 * `onOpen` is a reserved name: Apps Script runs it on open
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("SpiritHub")
    .addItem("Setup SpiritHub", "setup")
    .addSeparator()
    .addItem("Refresh Tournaments", "refreshTournaments")
    .addItem("Refresh Results", "refreshResults")
    .addItem("Refresh Issues", "refreshIssues")
    .addItem("Refresh Club Statistics", "refreshClubStatistics")
    .addToUi();
}

/**
 * rebuild the Clubs tab when a person changes something that affects club names
 * `onEdit` is a reserved name, apps script runs it after every edit by a person
 *
 * only reacts to the Club Override column on Teams and anything on Name Rules
 * skips quietly if another run holds the lock
 *   Refresh Results rebuilds Clubs at the end anyway
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e  the edit event
 */
function onEdit(e) {
  const sheetName = e.range.getSheet().getName();
  const overrideColumn = TEAM_KEYS.indexOf("clubOverride") + 1;
  const touchesOverride = sheetName === TEAMS_SHEET
    && e.range.getColumn() <= overrideColumn
    && e.range.getLastColumn() >= overrideColumn;
  if (!touchesOverride && sheetName !== NAME_RULES_SHEET) return;

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(1000)) return;
  try {
    _rebuildClubs_();
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
 * @template T
 * @param {function(): T} work the function to run
 * @returns {T} whatever `work` returns
 * @throws {Error} if another run still holds the lock after LOCK_WAIT_MS
 */
function _withLock_(work) {
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
 * @param {string} message text to show
 */
function _notify_(message) {
  console.log(message);
  SpreadsheetApp.getActiveSpreadsheet().toast(message, "SpiritHub", 10);
}

/**
 * run a function and log how long it took, to find slow steps
 * shows in the Apps Script Executions log
 *
 * @template T
 * @param {string}      label  name of the step
 * @param {function(): T} work the step
 * @returns {T} whatever `work` returns
 */
function _timed_(label, work) {
  const start = Date.now();
  try {
    return work();
  } finally {
    console.log(`${label} ${((Date.now() - start) / 1000).toFixed(1)}s`);
  }
}

/**
 * menu: scan the category folder and update the Tournaments tab
 */
function refreshTournaments() {
  _notify_(_withLock_(_refreshTournaments_));
}

/**
 * menu: import every tournament marked NEW or REFRESH, then refresh issues
 */
function refreshResults() {
  _notify_(_withLock_(_refreshResults_));
}

/**
 * menu: add any new issues from what is already in the spreadsheet
 */
function refreshIssues() {
  _notify_(_withLock_(() => _timed_("issues", _refreshIssues_)));
}

/**
 * menu: recalculate the Club Statistics tab from what is already in the spreadsheet
 * only ever run from the menu
 */
function refreshClubStatistics() {
  _notify_(_withLock_(_refreshClubStatistics_));
}
