/**
 * add the sotg hub menu when the spreadsheet is opened
 *
 * `onOpen` is a reserved name: Apps Script runs it on open
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("SOTG Hub")
    .addItem("Scan for files", "scanEvents")
    .addItem("Import new and refreshed events", "importEvents")
    .addToUi();
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
