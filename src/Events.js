/**
 * name of the Events tab
 */
const EVENTS_SHEET = "Events";

/**
 * text for the info row above the Events table
 * what the tab is, what to edit, how it refreshes
 */
const EVENTS_INFO = "All the found spirit results files in the folder for this category (e.g. University, Club)\n\n"
  + "User editable columns:\n"
  + "- Name Override: set a name for the tournament\n"
  + "- Status: filled on refresh, set to REFRESH to re-import results\n"
  + "- International: is this an international tournament\n"
  + "- Import: Should these results be inluded in spirit award and issue tracking\n\n"
  + "To Refresh run \"Import new and refreshed events\"";

/**
 * events table columns, in order
 * keys are the names used in code
 * values are the header text
 */
const EVENT_HEADERS = Object.freeze(
  /** @type {const} */ ({
    fileId: "File ID",
    fileName: "File Name",
    path: "Path",
    defaultName: "Default Name",
    nameOverride: "Name Override",
    status: "Status",
    message: "Message",
    international: "International",
    include: "Include",
    importedVersion: "Imported Version",
  }),
);

/**
 * values of the status column
 * OK is shown as a blank cell
 */
const EVENT_STATUS = Object.freeze(
  /** @type {const} */ ({
    NEW: "NEW",
    EDITED: "EDITED",
    REFRESH: "REFRESH",
    ERROR: "ERROR",
    MISSING: "MISSING",
    OK: "",
  }),
);

/**
 * One row of the Events tab.
 *
 * @typedef {Object} EventRecord
 * @property {string}    fileId           Drive file ID (the key).
 * @property {string}    fileName         File name.
 * @property {string}    path             Folder path below the category, joined with " / ".
 * @property {string}    defaultName      Tournament name used when there is no override (the folder name).
 * @property {string}    nameOverride     Tournament name typed by a person; "" if none.
 * @property {string}    status           One of EVENT_STATUS.
 * @property {string}    message          Explanation for ERROR / MISSING; "" otherwise.
 * @property {boolean}   international    Ticked for international events.
 * @property {boolean}   include          Ticked to include the event in calculations.
 * @property {Date|null} importedVersion  File's last-modified time when last imported; null if never imported.
 */

/**
 * keys of EventRecord in column order
 */
const EVENT_KEYS = /** @type {(keyof typeof EVENT_HEADERS)[]} */ (Object.keys(EVENT_HEADERS));

/**
 * changes smaller than this are ignored when comparing modified times
 */
const VERSION_TOLERANCE_MS = 1000;

/**
 * get the events tab, creating it with header row on first use.
 *
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} The Events tab.
 */
function _getEventsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  /** @type {readonly string[]} */
  const headers = Object.values(EVENT_HEADERS);
  return _getOrCreateSheet_(ss, EVENTS_SHEET, headers, EVENTS_INFO).sheet;
}

/**
 * convert one row of Events tab values into an EventRecord
 *
 * @param {unknown[]} row  values of one data row, in EVENT_KEYS order
 * @returns {EventRecord} the record
 */
function _eventFromRow_(row) {
  /** @param {keyof typeof EVENT_HEADERS} key */
  const cell = (key) => row[EVENT_KEYS.indexOf(key)];
  const imported = cell("importedVersion");
  return {
    fileId: String(cell("fileId")),
    fileName: String(cell("fileName")),
    path: String(cell("path")),
    defaultName: String(cell("defaultName")),
    nameOverride: String(cell("nameOverride")).trim(),
    status: String(cell("status")).trim().toUpperCase(),
    message: String(cell("message")),
    international: cell("international") === true,
    include: cell("include") === true,
    importedVersion: imported instanceof Date ? imported : null,
  };
}

/**
 * convert an EventRecord into a row of values for the Events tab
 *
 * @param {EventRecord} event  the record
 * @returns {unknown[]} values in EVENT_KEYS order
 */
function _eventToRow_(event) {
  return EVENT_KEYS.map((key) => {
    const value = event[key];
    return value === null ? "" : value;
  });
}

/**
 * read every data row of the Events tab
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet  the events tab
 * @returns {EventRecord[]} one record per data row
 */
function _readEvents_(sheet) {
  const rowCount = sheet.getLastRow() - HEADER_ROW;
  if (rowCount < 1) return [];
  return sheet
    .getRange(DATA_ROW, 1, rowCount, EVENT_KEYS.length)
    .getValues()
    .map(_eventFromRow_)
    .filter((event) => event.fileId !== "");
}

/**
 * write events to the events tab, starting at row 2
 * set the Status dropdown and tick boxes on exactly those rows
 *
 * Validation is only applied to rows that hold an event: tick boxes store FALSE,
 * so applying them to empty rows would make those rows look like data.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet   The Events tab.
 * @param {EventRecord[]}                      events  Records to write, in display order.
 */
function _writeEvents_(sheet, events) {
  _fitSheet_(sheet, HEADER_ROW + events.length, EVENT_KEYS.length);
  if (events.length === 0) return;
  sheet.getRange(DATA_ROW, 1, events.length, EVENT_KEYS.length).setValues(events.map(_eventToRow_));

  /** @param {keyof typeof EVENT_HEADERS} key */
  const column = (key) => EVENT_KEYS.indexOf(key) + 1;

  const statuses = Object.values(EVENT_STATUS).filter((s) => s !== "");
  const statusRule = SpreadsheetApp.newDataValidation().requireValueInList(statuses, true).build();
  sheet.getRange(DATA_ROW, column("status"), events.length, 1).setDataValidation(statusRule);

  const checkbox = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  sheet.getRange(DATA_ROW, column("international"), events.length, 1).setDataValidation(checkbox);
  sheet.getRange(DATA_ROW, column("include"), events.length, 1).setDataValidation(checkbox);
}

/**
 * whether a file has changed since it was last imported
 *
 * @param {Date|null} importedVersion  modified time recorded at import, null if never imported
 * @param {Date}      lastUpdated      file's current modified time
 * @returns {boolean} if the file is newer than the imported version
 */
function _isChangedSince_(importedVersion, lastUpdated) {
  return importedVersion !== null && lastUpdated.getTime() > importedVersion.getTime() + VERSION_TOLERANCE_MS;
}

/**
 * work out the new Events list from the current rows and a fresh scan
 *
 * existing rows keep their order and human-entered columns
 * new files are appended as NEW
 * files no longer found are marked MISSING
 *
 * @param {EventRecord[]} existing  current rows of the Events tab
 * @param {ResultsFile[]} files    results files found by the scan
 * @returns {EventRecord[]} updated list of events
 */
function _syncEvents_(existing, files) {
  const fileById = new Map(files.map((f) => [f.id, f]));
  const knownIds = new Set(existing.map((e) => e.fileId));

  const updated = existing.map((event) => {
    const file = fileById.get(event.fileId);
    if (!file) {
      return { ...event, status: EVENT_STATUS.MISSING, message: "File is no longer in the category folder" };
    }

    const changed = _isChangedSince_(event.importedVersion, file.lastUpdated);
    let status = event.status;
    let message = event.message;
    if (status === EVENT_STATUS.MISSING) {
      status = event.importedVersion === null ? EVENT_STATUS.NEW : changed ? EVENT_STATUS.EDITED : EVENT_STATUS.OK;
      message = "";
    } else if (status === EVENT_STATUS.OK && changed) {
      status = EVENT_STATUS.EDITED;
    }

    return {
      ...event,
      fileName: file.name,
      folderId: file.folderId,
      path: file.path.join(" / "),
      defaultName: file.folderName,
      status: status,
      message: message,
    };
  });

  const added = files
    .filter((f) => !knownIds.has(f.id))
    .map((f) => ({
      fileId: f.id,
      fileName: f.name,
      folderId: f.folderId,
      path: f.path.join(" / "),
      defaultName: f.folderName,
      nameOverride: "",
      status: EVENT_STATUS.NEW,
      message: "",
      international: false,
      include: true,
      importedVersion: null,
    }));

  return [...updated, ...added];
}

/**
 * the tournament name to show: the override if set, otherwise the default name
 *
 * @param {EventRecord} event  the event
 * @returns {string} display name
 */
function _tournamentName_(event) {
  return event.nameOverride || event.defaultName;
}

/**
 * scan the category folder and bring the Events tab up to date
 *
 * Does not import any responses
 *
 * @returns {string} a one-line summary of the Events tab
 */
function _scanEvents_() {
  const sheet = _getEventsSheet_();
  const events = _syncEvents_(_readEvents_(sheet), _scanCategory_(_readConfig_().category));
  _writeEvents_(sheet, events);

  /** @type {Record<string, number>} */
  const counts = {};
  for (const e of events) {
    const label = e.status || "up to date";
    counts[label] = (counts[label] || 0) + 1;
  }
  const breakdown = Object.entries(counts).map(([status, n]) => `${n} ${status}`).join(", ");
  return `${events.length} events: ${breakdown}`;
}
