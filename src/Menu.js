/**
 * add the sotg hub menu when the spreadsheet is opened
 *
 * `onOpen` is a reserved name: Apps Script runs it on open
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("SOTG Hub")
    .addItem("Setup hub", "setup")
    .addSeparator()
    .addItem("Scan for files", "scanEvents")
    .addItem("Import new and refreshed events", "importEvents")
    .addToUi();
}

/**
 * rebuild the Clubs tab when a person changes something that affects club names
 * `onEdit` is a reserved name, apps script runs it after every edit by a person
 *
 * only reacts to the Club Override column on Teams and anything on Name Rules
 * skips quietly if another run holds the lock
 *   an import rebuilds Clubs at the end anyway
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
 * how long to wait for another run to finish be giving up
 */
const LOCK_WAIT_MS = 1000;

/**
 * run a function while holding this spreadsheet's sript lock
 * so two runs (by same or different people) can never conflict
 * @template T
 * @param {function(): T} work the function to run
 * @returns {T} whatever `work` returns
 * @throws {Error} if another run still holds the lock after LOCK_WAIT_MS
 */
function _withLock_(work) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(LOCK_WAIT_MS)) {
    throw new Error("Another SOTG Hub run is in progress. Please try again in a minute");
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
  SpreadsheetApp.getActiveSpreadsheet().toast(message, "SOTG Hub", 10);
}

/**
 * Menu: scan the category folder and update the Events tab.
 */
function scanEvents() {
  _notify_(_withLock_(_scanEvents_));
}

/**
 * Menu: import every event marked NEW or REFRESH.
 */
function importEvents() {
  _notify_(_withLock_(_importEvents_));
}
