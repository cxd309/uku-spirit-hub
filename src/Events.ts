/**
 * name of the Events tab
 */
const EVENTS_SHEET = "Tournaments";

/**
 * text for the info row above the Events table
 * what the tab is, what to edit, how it refreshes
 */
const EVENTS_INFO = "All the found spirit results files in the folder for this category (e.g. University, Club)\n"
  + "Each results file must be in a folder named \"YYYYMMDD Tournament name\", rename the folder to rename the tournament\n\n"
  + "User editable columns:\n"
  + "- Status: filled on refresh, set to REFRESH to re-import results\n"
  + "- International: is this an international tournament\n"
  + "- Include: should these results be included in the spirit award and issue tracking\n\n"
  + "To find new or edited files run \"Refresh Tournaments\", to import them run \"Refresh Results\"";

/**
 * events table columns, in order
 * keys are the names used in code
 * values are the header text
 */
const EVENT_HEADERS = Object.freeze(
  {
    fileId: "File ID",
    fileName: "File Name",
    path: "Path",
    date: "Date",
    tournament: "Tournament",
    status: "Status",
    message: "Message",
    international: "International",
    include: "Include",
    importedVersion: "Imported Version",
  } as const,
);

/**
 * values of the status column
 * OK is shown as a blank cell
 */
const EVENT_STATUS = Object.freeze(
  {
    NEW: "NEW",
    EDITED: "EDITED",
    REFRESH: "REFRESH",
    ERROR: "ERROR",
    MISSING: "MISSING",
    OK: "",
  } as const,
);

/**
 * One row of the Events tab.
 */
interface EventRecord {
  /** Drive file ID (the key). */
  fileId: string;
  /** File name. */
  fileName: string;
  /** Folder path below the category, joined with " / ". */
  path: string;
  /** tournament date from the folder name, null if the folder name is invalid */
  date: Date | null;
  /** tournament name from the folder name, "" if the folder name is invalid */
  tournament: string;
  /** One of EVENT_STATUS. */
  status: string;
  /** Explanation for ERROR / MISSING; "" otherwise. */
  message: string;
  /** Ticked for international events. */
  international: boolean;
  /** Ticked to include the event in calculations. */
  include: boolean;
  /** File's last-modified time when last imported; null if never imported. */
  importedVersion: Date | null;
}

/**
 * keys of EventRecord in column order
 */
const EVENT_KEYS = Object.keys(EVENT_HEADERS) as (keyof typeof EVENT_HEADERS)[];

/**
 * changes smaller than this are ignored when comparing modified times
 */
const VERSION_TOLERANCE_MS = 1000;

/**
 * get the events tab, creating it with header row on first use.
 *
 * @returns The Events tab.
 */
function _getEventsSheet_(): GoogleAppsScript.Spreadsheet.Sheet {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const headers: readonly string[] = Object.values(EVENT_HEADERS);
  return _getOrCreateSheet_(ss, EVENTS_SHEET, headers, EVENTS_INFO);
}

/**
 * convert one row of Events tab values into an EventRecord
 *
 * @param row  values of one data row, in EVENT_KEYS order
 * @returns the record
 */
function _eventFromRow_(row: unknown[]): EventRecord {
  const cell = (key: keyof typeof EVENT_HEADERS) => row[EVENT_KEYS.indexOf(key)];
  const imported = cell("importedVersion");
  return {
    fileId: String(cell("fileId")),
    fileName: String(cell("fileName")),
    path: String(cell("path")),
    date: cell("date") instanceof Date ? cell("date") as Date : null,
    tournament: String(cell("tournament")),
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
 * @param event  the record
 * @returns values in EVENT_KEYS order
 */
function _eventToRow_(event: EventRecord): unknown[] {
  return EVENT_KEYS.map((key) => {
    const value = event[key];
    return value === null ? "" : value;
  });
}

/**
 * read every data row of the Events tab
 *
 * @param sheet  the events tab
 * @returns one record per data row
 */
function _readEvents_(sheet: GoogleAppsScript.Spreadsheet.Sheet): EventRecord[] {
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
 * @param sheet   The Events tab.
 * @param events  Records to write, in display order.
 */
function _writeEvents_(sheet: GoogleAppsScript.Spreadsheet.Sheet, events: EventRecord[]) {
  _writeTable_(sheet, events.map(_eventToRow_), EVENT_KEYS.length);
  if (events.length === 0) return;

  const column = (key: keyof typeof EVENT_HEADERS) => EVENT_KEYS.indexOf(key) + 1;

  const statuses = Object.values(EVENT_STATUS).filter((s) => s !== "");
  const statusRule = SpreadsheetApp.newDataValidation().requireValueInList(statuses, true).build();
  sheet.getRange(DATA_ROW, column("status"), events.length, 1).setDataValidation(statusRule);

  const checkbox = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  sheet.getRange(DATA_ROW, column("international"), events.length, 1).setDataValidation(checkbox);
  sheet.getRange(DATA_ROW, column("include"), events.length, 1).setDataValidation(checkbox);

  sheet.getRange(DATA_ROW, column("date"), events.length, 1).setNumberFormat("yyyy-mm-dd");
  _highlightDuplicateTournaments_(sheet, events.length);
}

/**
 * whether a file has changed since it was last imported
 *
 * @param importedVersion  modified time recorded at import, null if never imported
 * @param lastUpdated      file's current modified time
 * @returns if the file is newer than the imported version
 */
function _isChangedSince_(importedVersion: Date | null, lastUpdated: Date): boolean {
  return importedVersion !== null && lastUpdated.getTime() > importedVersion.getTime() + VERSION_TOLERANCE_MS;
}

/**
 * work out the new Events list from the current rows and a fresh scan
 *
 * existing rows keep their order and human-entered columns
 * new files are appended as NEW
 * files no longer found are marked MISSING
 *
 * @param existing  current rows of the Events tab
 * @param files     results files found by the scan
 * @returns updated list of events
 */
function _syncEvents_(existing: EventRecord[], files: ResultsFile[]): EventRecord[] {
  const fileById = new Map(files.map((f) => [f.id, f]));
  const knownIds = new Set(existing.map((e) => e.fileId));

  const updated = existing.map((event) => {
    const file = fileById.get(event.fileId);
    if (!file) {
      return { ...event, status: EVENT_STATUS.MISSING, message: "File is no longer in the category folder" };
    }

    const changed = _isChangedSince_(event.importedVersion, file.lastUpdated);
    const problem = _fileProblem_(file.name, file.folderName);
    let status = event.status;
    let message = event.message;
    if (problem !== "") {
      status = EVENT_STATUS.ERROR;
      message = problem;
    } else if (status === EVENT_STATUS.MISSING || (status === EVENT_STATUS.ERROR && _isFileProblem_(message))) {
      // the file is back, or its name has been fixed
      status = event.importedVersion === null ? EVENT_STATUS.NEW : changed ? EVENT_STATUS.EDITED : EVENT_STATUS.OK;
      message = "";
    } else if (status === EVENT_STATUS.OK && changed) {
      status = EVENT_STATUS.EDITED;
    }

    return {
      ...event,
      fileName: file.name,
      path: file.path.join(" / "),
      ..._folderFields_(file.folderName),
      status: status,
      message: message,
    };
  });

  const added = files
    .filter((f) => !knownIds.has(f.id))
    .map((f) => ({
      fileId: f.id,
      fileName: f.name,
      path: f.path.join(" / "),
      ..._folderFields_(f.folderName),
      status: _fileProblem_(f.name, f.folderName) === "" ? EVENT_STATUS.NEW : EVENT_STATUS.ERROR,
      message: _fileProblem_(f.name, f.folderName),
      international: false,
      include: true,
      importedVersion: null,
    }));

  return _sortEvents_([...updated, ...added]);
}

/**
 * pattern every event folder name must match
 * YYYYMMDD, one or more spaces, then the tournament name
 */
const FOLDER_NAME_PATTERN = /^(\d{4})(\d{2})(\d{2})\s+(.+)$/;

/**
 * result of reading an event folder name
 */
type FolderNameResult = { ok: true; date: Date; tournament: string } | { ok: false; reason: string };

/**
 * read the date and tournament name from an event folder name
 * pure, no google calls
 * the date must be a real calendar date, e.g. 20260000 is rejected
 *
 * @param name  folder name, e.g. "20251101 ELUXIR"
 * @returns date and tournament, or why the name is invalid
 */
function _parseFolderName_(name: string): FolderNameResult {
  const match = FOLDER_NAME_PATTERN.exec(name.trim());
  if (!match) {
    return { ok: false, reason: `folder name "${name}" must be "YYYYMMDD Tournament name"` };
  }
  const [, year, month, day, tournament] = match;
  const date = new Date(Number(year), Number(month) - 1, Number(day));
  if (date.getFullYear() !== Number(year) || date.getMonth() !== Number(month) - 1 || date.getDate() !== Number(day)) {
    return { ok: false, reason: `folder name "${name}" starts with ${year}${month}${day}, which is not a real date` };
  }
  return { ok: true, date: date, tournament: tournament.trim() };
}

/**
 * why a file found by the scan cannot be imported, from its file and folder names
 * pure, no google calls
 * every reason starts with "file name" or "folder name", see _isFileProblem_
 *
 * @param fileName    name of the file
 * @param folderName  name of the folder containing the file
 * @returns the reason, "" if the names are fine
 */
function _fileProblem_(fileName: string, folderName: string): string {
  if (!fileName.toLowerCase().includes(RESULTS_FILE_TEXT.toLowerCase())) {
    return `file name "${fileName}" does not contain "${RESULTS_FILE_TEXT}"`;
  }
  const folder = _parseFolderName_(folderName);
  return folder.ok ? "" : folder.reason;
}

/**
 * whether a Message was written because of a bad file or folder name
 * these are cleared by the next Refresh Tournaments once the name is fixed
 * other errors (e.g. a failed import) stay until a person sets REFRESH
 *
 * @param message  the Message cell
 * @returns true if it came from _fileProblem_ or the folder name check on import
 */
function _isFileProblem_(message: string): boolean {
  return message.startsWith("file name ") || message.startsWith("folder name ");
}

/**
 * date and tournament columns for an event, from its folder name
 * both blank when the folder name is invalid
 *
 * @param folderName  name of the folder containing the results file
 * @returns values for the Date and Tournament columns
 */
function _folderFields_(folderName: string): { date: Date | null; tournament: string } {
  const folder = _parseFolderName_(folderName);
  return folder.ok ? { date: folder.date, tournament: folder.tournament } : { date: null, tournament: "" };
}

/**
 * highlight tournament names that appear more than once in red
 * blank names are highlighted too: the folder name could not be read, so the tournament cannot be imported
 * replaces only rules on the Tournament column, other conditional formatting is kept
 * matching is case-insensitive, only event rows are covered so empty rows below the table are never highlighted
 *
 * @param sheet  the Events tab
 * @param rows   number of event rows
 */
function _highlightDuplicateTournaments_(sheet: GoogleAppsScript.Spreadsheet.Sheet, rows: number) {
  const column = EVENT_KEYS.indexOf("tournament") + 1;
  const letter = _columnLetter_(column);
  const range = sheet.getRange(DATA_ROW, column, rows, 1);
  const rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(
      `=OR($${letter}${DATA_ROW}="", COUNTIF($${letter}$${DATA_ROW}:$${letter}, $${letter}${DATA_ROW})>1)`,
    )
    .setBackground("#f4c7c3")
    .setFontColor("#a50e0e")
    .setRanges([range])
    .build();
  const others = sheet
    .getConditionalFormatRules()
    .filter((r) => !r.getRanges().some((rg) => rg.getColumn() === column));
  sheet.setConditionalFormatRules([...others, rule]);
}

/**
 * scan the category folder and bring the Events tab up to date
 *
 * does not import any responses, Refresh Results does that
 *
 * @returns a one-line summary of the Events tab
 */
function _refreshTournaments_(): string {
  const sheet = _getEventsSheet_();
  const category = _readHubSettings_().category;
  const existing = _readEvents_(sheet);
  _checkCategory_(category, existing.length);
  const files = _timed_("scan drive", () => _scanCategory_(category));
  const events = _syncEvents_(existing, files);
  _timed_("write tournaments", () => _writeEvents_(sheet, events));
  _rememberCategory_(category);

  // tournament names come from folder names, so Teams and Clubs pick up any renames
  const teamsSheet = _getTeamsSheet_();
  const responses = _readResponses_(_getResponsesSheet_());
  _timed_("write teams", () => _writeTeams_(teamsSheet, _readTeams_(teamsSheet), responses, events));
  _timed_("rebuild clubs", () => {
    SpreadsheetApp.flush();
    _rebuildClubs_();
  });

  const counts: Record<string, number> = {};
  for (const e of events) {
    const label = e.status || "up to date";
    counts[label] = (counts[label] || 0) + 1;
  }
  const breakdown = Object.entries(counts).map(([status, n]) => `${n} ${status}`).join(", ");
  return `${events.length} events: ${breakdown}`;
}

/**
 * events in date order, oldest first
 * then by tournament and file name, so events on the same day have a fixed order
 * pure, returns a new array
 *
 * @param events  events in any order
 * @returns events in date order
 */
function _sortEvents_(events: EventRecord[]): EventRecord[] {
  return [...events].sort((a, b) =>
    _compareDates_(a.date, b.date)
    || a.tournament.localeCompare(b.tournament)
    || a.fileName.localeCompare(b.fileName)
  );
}
