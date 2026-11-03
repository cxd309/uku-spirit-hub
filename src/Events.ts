/**
 * name of the Events tab
 */
const EVENTS_SHEET = "Tournaments";

/**
 * text for the info row above the Events table
 * what the tab is, what to edit, how it refreshes
 */
const EVENTS_INFO =
  "Every tournament on the Schedule List of the Schedule Manager within the dates on the Settings tab\n"
  + "Rows without a Schedule URL or SOTG Sheet are left out, the SOTG Sheet link identifies each tournament\n\n"
  + "Grey columns are filled on refresh, white columns are yours to edit:\n"
  + "- Event Name Override: a name to use instead of the Schedule Manager's, leave blank to use theirs\n"
  + "- Schedule Status / Spirit Form Status: set REFRESH to import that file again\n"
  + "- Include Results: import this tournament's results, unticked rows are NOT INCLUDED\n"
  + "- International: only the receiving teams are UK teams\n"
  + "  scoring teams are left off Teams and Clubs, and missing spirit scores are not checked\n\n"
  + "Statuses: NEW not imported yet, EDITED changed since import, REFRESH import again, ERROR cannot be read, "
  + "MISSING no longer on the Schedule List, OK up to date\n"
  + "Event Status is the first of NOT INCLUDED, MISSING, ERROR, REFRESH, EDITED, NEW that either file has, "
  + "otherwise OK\n\n"
  + "To read the Schedule Manager again run \"Refresh Tournaments\", to import results run \"Refresh Results\"";

/**
 * events table columns, in order
 * keys are the names used in code
 * values are the header text
 * columns listed in EVENT_FORMULAS are formulas, the rest hold EventRecord values
 */
const EVENT_HEADERS = Object.freeze(
  {
    scheduleEventName: "Schedule Event Name",
    eventNameOverride: "Event Name Override",
    eventName: "Event Name",
    startDate: "Start Date",
    endDate: "End Date",
    scheduleUrl: "Schedule URL",
    scheduleImportDate: "Schedule Import Date",
    scheduleStatus: "Schedule Status",
    spiritFormUrl: "Spirit Form URL",
    spiritFormImportDate: "Spirit Form Import Date",
    spiritFormStatus: "Spirit Form Status",
    eventStatus: "Event Status",
    message: "Message",
    includeResults: "Include Results",
    international: "International",
  } as const,
);

/**
 * keys of EventRecord in column order
 */
const EVENT_KEYS = Object.keys(EVENT_HEADERS) as (keyof typeof EVENT_HEADERS)[];

/**
 * values of the Schedule Status and Spirit Form Status columns
 */
const SOURCE_STATUS = Object.freeze(
  {
    NEW: "NEW",
    EDITED: "EDITED",
    REFRESH: "REFRESH",
    ERROR: "ERROR",
    MISSING: "MISSING",
    OK: "OK",
  } as const,
);

/**
 * Event Status for a tournament with Include Results unticked
 */
const NOT_INCLUDED = "NOT INCLUDED";

/**
 * file statuses in the order they decide Event Status: the first either file has wins, otherwise OK
 */
const EVENT_STATUS_ORDER = Object.freeze(
  [
    SOURCE_STATUS.MISSING,
    SOURCE_STATUS.ERROR,
    SOURCE_STATUS.REFRESH,
    SOURCE_STATUS.EDITED,
    SOURCE_STATUS.NEW,
  ] as const,
);

/**
 * columns a person edits, left white, every other column is shaded as generated
 */
const EVENT_EDITABLE = Object.freeze(
  ["eventNameOverride", "scheduleStatus", "spiritFormStatus", "includeResults", "international"] as const,
);

/**
 * background of the generated columns, light grey 1 as on the Schedule Manager
 */
const GENERATED_BACKGROUND = "#d9d9d9";

/**
 * tick box columns of the Events tab
 */
const EVENT_TOGGLES = Object.freeze(["includeResults", "international"] as const);

/**
 * changes smaller than this are ignored when comparing modified times
 */
const VERSION_TOLERANCE_MS = 1000;

/**
 * One row of the Events tab.
 */
interface EventRecord {
  /** Event Name on the Schedule List */
  scheduleEventName: string;
  /** name typed in to use instead, "" to use the Schedule List's */
  eventNameOverride: string;
  /** the name used everywhere else: the override if there is one */
  eventName: string;
  /** Start date on the Schedule List, null if it could not be read */
  startDate: Date | null;
  /** End date on the Schedule List, null if blank */
  endDate: Date | null;
  /** link to the schedule spreadsheet */
  scheduleUrl: string;
  /** schedule's modified time when it was last imported, null if never imported */
  scheduleImportDate: Date | null;
  /** one of SOURCE_STATUS */
  scheduleStatus: string;
  /** link to the Form Responses spreadsheet (the key) */
  spiritFormUrl: string;
  /** Form Responses' modified time when it was last imported, null if never imported */
  spiritFormImportDate: Date | null;
  /** one of SOURCE_STATUS */
  spiritFormStatus: string;
  /** NOT_INCLUDED or the first of EVENT_STATUS_ORDER either file has, otherwise OK */
  eventStatus: string;
  /** why a file is ERROR or MISSING, "" otherwise */
  message: string;
  /** ticked to import this tournament's results */
  includeResults: boolean;
  /** ticked for international events */
  international: boolean;
}

/**
 * formula columns of the Events tab
 * each row gets its own formula, so sorting the tab never breaks them
 */
const EVENT_FORMULAS: Readonly<Partial<Record<keyof typeof EVENT_HEADERS, (row: number) => string>>> = Object.freeze({
  eventName: (row: number) => {
    const cell = (key: keyof typeof EVENT_HEADERS) => `$${_columnLetter_(EVENT_KEYS.indexOf(key) + 1)}${row}`;
    return `=IF(TRIM(${cell("eventNameOverride")})="", ${cell("scheduleEventName")}, TRIM(${
      cell("eventNameOverride")
    }))`;
  },
  eventStatus: (row: number) => {
    const cell = (key: keyof typeof EVENT_HEADERS) => `$${_columnLetter_(EVENT_KEYS.indexOf(key) + 1)}${row}`;
    const either = (status: string) =>
      `OR(${cell("scheduleStatus")}="${status}", ${cell("spiritFormStatus")}="${status}")`;
    const rest = EVENT_STATUS_ORDER.reduceRight(
      (otherwise, status) => `IF(${either(status)}, "${status}", ${otherwise})`,
      `"${SOURCE_STATUS.OK}"`,
    );
    return `=IF(NOT(${cell("includeResults")}), "${NOT_INCLUDED}", ${rest})`;
  },
});

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
 * the name a tournament is known by: the override if there is one, otherwise the Schedule List's
 * pure, the same as the Event Name formula
 *
 * @param scheduleEventName  Event Name on the Schedule List
 * @param override           Event Name Override
 * @returns the name
 */
function _eventName_(scheduleEventName: string, override: string): string {
  return override.trim() === "" ? scheduleEventName : override.trim();
}

/**
 * Event Status of a tournament, the same as the Event Status formula
 * pure, no google calls
 *
 * @param event  the tournament
 * @returns NOT_INCLUDED, the first of EVENT_STATUS_ORDER either file has, otherwise OK
 */
function _eventStatus_(event: Pick<EventRecord, "includeResults" | "scheduleStatus" | "spiritFormStatus">): string {
  if (!event.includeResults) return NOT_INCLUDED;
  return EVENT_STATUS_ORDER.find((s) => event.scheduleStatus === s || event.spiritFormStatus === s)
    ?? SOURCE_STATUS.OK;
}

/**
 * convert one row of Events tab values into an EventRecord
 * Event Name and Event Status are worked out again rather than read, so they never wait for the formulas
 *
 * @param row  values of one data row, in EVENT_KEYS order
 * @returns the record
 */
function _eventFromRow_(row: unknown[]): EventRecord {
  const cell = (key: keyof typeof EVENT_HEADERS) => row[EVENT_KEYS.indexOf(key)];
  const date = (key: keyof typeof EVENT_HEADERS) => (cell(key) instanceof Date ? cell(key) as Date : null);
  const status = (key: keyof typeof EVENT_HEADERS) => String(cell(key)).trim().toUpperCase();
  const event = {
    scheduleEventName: String(cell("scheduleEventName")),
    eventNameOverride: String(cell("eventNameOverride")),
    startDate: date("startDate"),
    endDate: date("endDate"),
    scheduleUrl: String(cell("scheduleUrl")).trim(),
    scheduleImportDate: date("scheduleImportDate"),
    scheduleStatus: status("scheduleStatus"),
    spiritFormUrl: String(cell("spiritFormUrl")).trim(),
    spiritFormImportDate: date("spiritFormImportDate"),
    spiritFormStatus: status("spiritFormStatus"),
    message: String(cell("message")),
    includeResults: cell("includeResults") === true,
    international: cell("international") === true,
  };
  return {
    ...event,
    eventName: _eventName_(event.scheduleEventName, event.eventNameOverride),
    eventStatus: _eventStatus_(event),
  };
}

/**
 * convert an EventRecord into a row of values for the Events tab
 *
 * @param event  the record
 * @returns values in EVENT_KEYS order, "" in formula columns (filled by _fillFormulaColumns_)
 */
function _eventToRow_(event: EventRecord): unknown[] {
  return EVENT_KEYS.map((key) => {
    if (EVENT_FORMULAS[key]) return "";
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
    .filter((event) => event.spiritFormUrl !== "");
}

/**
 * write events to the events tab, oldest first
 * set the status dropdowns, tick boxes, date formats and shading on exactly those rows
 *
 * Validation is only applied to rows that hold an event: tick boxes store FALSE,
 * so applying them to empty rows would make those rows look like data.
 *
 * @param sheet   The Events tab.
 * @param events  Records to write, in any order.
 */
function _writeEvents_(sheet: GoogleAppsScript.Spreadsheet.Sheet, events: EventRecord[]) {
  const rows = _sortEvents_(events);
  _writeTable_(sheet, rows.map(_eventToRow_), EVENT_KEYS.length);
  if (rows.length === 0) return;
  _fillFormulaColumns_(sheet, EVENT_KEYS, EVENT_FORMULAS, rows.length);

  const column = (key: keyof typeof EVENT_HEADERS) => EVENT_KEYS.indexOf(key) + 1;
  const range = (key: keyof typeof EVENT_HEADERS) => sheet.getRange(DATA_ROW, column(key), rows.length, 1);

  const statusRule = SpreadsheetApp.newDataValidation().requireValueInList(Object.values(SOURCE_STATUS), true).build();
  range("scheduleStatus").setDataValidation(statusRule);
  range("spiritFormStatus").setDataValidation(statusRule);

  const checkbox = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  for (const key of EVENT_TOGGLES) range(key).setDataValidation(checkbox);

  range("startDate").setNumberFormat("yyyy-mm-dd");
  range("endDate").setNumberFormat("yyyy-mm-dd");
  range("scheduleImportDate").setNumberFormat("yyyy-mm-dd hh:mm");
  range("spiritFormImportDate").setNumberFormat("yyyy-mm-dd hh:mm");

  const editable: readonly string[] = EVENT_EDITABLE;
  const shading = EVENT_KEYS.map((key) => (editable.includes(key) ? null : GENERATED_BACKGROUND));
  sheet.getRange(DATA_ROW, 1, rows.length, EVENT_KEYS.length).setBackgrounds(rows.map(() => shading));

  _highlightDuplicateTournaments_(sheet, rows.length);
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
 * the status of one file of a tournament, and why if it is not fine
 * pure, no google calls
 *
 * a file that cannot be read is ERROR, REFRESH set by a person stays until it is imported
 * otherwise NEW if never imported, EDITED if changed since, OK if not
 *
 * @param previous    the status on the tab now, "" for a new tournament
 * @param importDate  modified time when last imported, null if never imported
 * @param version     the file's modified time now, or why it could not be read
 * @returns the status and a reason, "" unless ERROR
 */
function _sourceStatus_(
  previous: string,
  importDate: Date | null,
  version: FileVersion | undefined,
): { status: string; reason: string } {
  if (!version) return { status: SOURCE_STATUS.ERROR, reason: "not checked" };
  if (!version.ok) return { status: SOURCE_STATUS.ERROR, reason: version.reason };
  if (previous === SOURCE_STATUS.REFRESH) return { status: SOURCE_STATUS.REFRESH, reason: "" };
  if (importDate === null) return { status: SOURCE_STATUS.NEW, reason: "" };
  if (_isChangedSince_(importDate, version.version)) return { status: SOURCE_STATUS.EDITED, reason: "" };
  return { status: SOURCE_STATUS.OK, reason: "" };
}

/**
 * work out the new Events list from the current rows and the Schedule List
 * pure, no google calls
 *
 * tournaments are matched by their Spirit Form URL, keeping the columns a person edits
 * names, dates and links always come from the Schedule List
 * a changed Schedule URL counts as a new schedule, so it is imported again
 * tournaments no longer listed are dropped if nothing was imported from them, otherwise marked MISSING
 * a Spirit Form URL listed more than once uses its first row
 *
 * @param existing  current rows of the Events tab
 * @param listed    tournaments on the Schedule List within the date range
 * @param versions  link → the file's modified time now, for every link in `listed`
 * @returns updated list of events
 */
function _syncEvents_(
  existing: EventRecord[],
  listed: ScheduleListRow[],
  versions: Map<string, FileVersion>,
): EventRecord[] {
  const existingByUrl = new Map(existing.map((e) => [e.spiritFormUrl, e]));
  const rowsByUrl: Map<string, ScheduleListRow[]> = new Map();
  for (const row of listed) rowsByUrl.set(row.spiritFormUrl, [...(rowsByUrl.get(row.spiritFormUrl) ?? []), row]);

  const updated = [...rowsByUrl].map(([url, rows]): EventRecord => {
    const row = rows[0];
    const previous = existingByUrl.get(url);
    const scheduleImportDate = previous && previous.scheduleUrl === row.scheduleUrl
      ? previous.scheduleImportDate
      : null;
    const spiritFormImportDate = previous?.spiritFormImportDate ?? null;
    const schedule = _sourceStatus_(
      previous && previous.scheduleUrl === row.scheduleUrl ? previous.scheduleStatus : "",
      scheduleImportDate,
      versions.get(row.scheduleUrl),
    );
    const spiritForm = _sourceStatus_(previous?.spiritFormStatus ?? "", spiritFormImportDate, versions.get(url));

    const messages: string[] = [];
    if (schedule.reason !== "") messages.push(`Schedule: ${schedule.reason}`);
    if (spiritForm.reason !== "") messages.push(`Spirit Form: ${spiritForm.reason}`);
    if (rows.length > 1) {
      messages.push(
        `SOTG Sheet is on rows ${
          rows.map((r) => r.sourceRow).join(", ")
        } of the Schedule List, row ${row.sourceRow} used`,
      );
    }

    const eventNameOverride = previous?.eventNameOverride ?? "";
    const event = {
      scheduleEventName: row.eventName,
      eventNameOverride,
      eventName: _eventName_(row.eventName, eventNameOverride),
      startDate: row.startDate,
      endDate: row.endDate,
      scheduleUrl: row.scheduleUrl,
      scheduleImportDate,
      scheduleStatus: schedule.status,
      spiritFormUrl: url,
      spiritFormImportDate,
      spiritFormStatus: spiritForm.status,
      message: messages.join("; "),
      includeResults: previous?.includeResults ?? false,
      international: previous?.international ?? false,
    };
    return { ...event, eventStatus: _eventStatus_(event) };
  });

  const gone = existing
    .filter((e) => !rowsByUrl.has(e.spiritFormUrl))
    .filter((e) => e.scheduleImportDate !== null || e.spiritFormImportDate !== null)
    .map((e): EventRecord => {
      const event = {
        ...e,
        scheduleStatus: SOURCE_STATUS.MISSING,
        spiritFormStatus: SOURCE_STATUS.MISSING,
        message: "No longer on the Schedule List within the dates on the Settings tab",
      };
      return { ...event, eventStatus: _eventStatus_(event) };
    });

  return [...updated, ...gone];
}

/**
 * highlight event names that appear more than once in red
 * blank names are highlighted too, a tournament needs a name to be told apart in Issues and the Club Report
 * replaces only rules on the Event Name column, other conditional formatting is kept
 * matching is case-insensitive, only event rows are covered so empty rows below the table are never highlighted
 *
 * @param sheet  the Events tab
 * @param rows   number of event rows
 */
function _highlightDuplicateTournaments_(sheet: GoogleAppsScript.Spreadsheet.Sheet, rows: number) {
  const column = EVENT_KEYS.indexOf("eventName") + 1;
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
 * read the Schedule List of the Schedule Manager and bring the Events tab up to date
 *
 * does not import any responses, Refresh Results does that
 *
 * @returns a one-line summary of the Events tab
 */
function _refreshTournaments_(): string {
  const settings = _readHubSettings_();
  const sheet = _getEventsSheet_();
  const existing = _readEvents_(sheet);
  const list = _timed_("read schedule list", () =>
    _parseScheduleList_(
      _readScheduleListValues_(settings.scheduleManagerUrl),
      settings.startDate,
      settings.endDate,
    ));
  const versions = _timed_(
    "read file versions",
    () => _readFileVersions_(list.rows.flatMap((r) => [r.scheduleUrl, r.spiritFormUrl])),
  );
  const events = _syncEvents_(existing, list.rows, versions);
  _timed_("write tournaments", () => _writeEvents_(sheet, events));

  // event names come from the Schedule List, so Teams and Clubs pick up any renames and International ticks
  _rebuildTeamsAndClubs_(_readResponseTeams_(_getResponsesSheet_()), _sortEvents_(events));

  const counts: Record<string, number> = {};
  for (const e of events) counts[e.eventStatus] = (counts[e.eventStatus] || 0) + 1;
  const breakdown = Object.entries(counts).map(([status, n]) => `${n} ${status}`).join(", ");
  const skipped = list.noStartDate.length === 0
    ? ""
    : `; Schedule List rows with no start date left out: ${list.noStartDate.join(", ")}`;
  return `${events.length} tournaments${events.length > 0 ? `: ${breakdown}` : ""}${skipped}`;
}

/**
 * events in date order, oldest first, the order of the Tournaments tab
 * then by event name and Spirit Form URL, so events on the same day have a fixed order
 * pure, returns a new array
 *
 * @param events  events in any order
 * @returns events in date order
 */
function _sortEvents_(events: EventRecord[]): EventRecord[] {
  return [...events].sort((a, b) =>
    _compareDates_(a.startDate, b.startDate)
    || a.eventName.localeCompare(b.eventName)
    || a.spiritFormUrl.localeCompare(b.spiritFormUrl)
  );
}
