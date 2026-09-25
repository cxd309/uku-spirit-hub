/**
 * menu: create every tab this hub needs, in order
 * tabs that already exist are left exactly as they are
 * safe to run any number of times
 */
function setup() {
  _notify_(_withLock_(_setup_));
}

/**
 * create any missing tabs, in the order they should appear
 * removes the empty default "Sheet1" once the hub tabs exist
 *
 * @returns {string} one-line summary
 */
function _setup_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const before = new Set(ss.getSheets().map((s) => s.getName()));

  _getEventsSheet_();
  _getResponsesSheet_();
  _getTeamsSheet_();
  _getClubsSheet_();
  _getNameRulesSheet_();
  _getConfigSheet_();

  const blank = ss.getSheetByName("Sheet1");
  if (blank && blank.getLastRow() === 0 && blank.getLastColumn() === 0) ss.deleteSheet(blank);

  const created = ss.getSheets().map((s) => s.getName()).filter((n) => !before.has(n));
  return created.length === 0 ? "Hub already set up: nothing to do" : `Created: ${created.join(", ")}`;
}
