/**
 * a results file fetched from Google: its modified time and every tab, or why it could not be fetched
 */
type FileRead = { ok: true; version: Date; tabs: ResultsTab[] } | { ok: false; reason: string };

/**
 * what is asked of the Sheets API for a results file: every tab's name and cell values, nothing else
 */
const RESULTS_FILE_FIELDS = "sheets(properties(title),data(rowData(values(effectiveValue,formattedValue))))";

/**
 * how long to wait before retrying requests Google asked to slow down
 */
const RETRY_WAIT_MS = 5000;

/**
 * send GET requests to Google APIs all at once, as the person running the script
 * requests Google answers with 429 (too many requests) are retried once, after RETRY_WAIT_MS
 *
 * @param urls  API urls
 * @returns each response's status code and parsed body, in the order given
 */
function _fetchGoogleJson_(urls: string[]): { code: number; body: unknown }[] {
  const headers = { Authorization: `Bearer ${ScriptApp.getOAuthToken()}` };
  const fetch = (list: string[]) =>
    UrlFetchApp.fetchAll(list.map((url) => ({ url, headers, muteHttpExceptions: true })))
      .map((response) => ({ code: response.getResponseCode(), text: response.getContentText() }));

  const results = fetch(urls);
  const retry = results.flatMap((r, i) => (r.code === 429 ? [i] : []));
  if (retry.length > 0) {
    Utilities.sleep(RETRY_WAIT_MS);
    fetch(retry.map((i) => urls[i])).forEach((r, j) => {
      results[retry[j]] = r;
    });
  }
  return results.map((r) => {
    try {
      return { code: r.code, body: JSON.parse(r.text) as unknown };
    } catch (e) {
      return { code: r.code, body: {} };
    }
  });
}

/**
 * the message from a failed Google API response
 *
 * @param response  status code and parsed body
 * @returns Google's error message, or the status code if there is none
 */
function _apiError_(response: { code: number; body: unknown }): string {
  const message = (response.body as { error?: { message?: string } } | null)?.error?.message;
  return message ?? `HTTP ${response.code}`;
}

/**
 * a cell as the Sheets API returns it, turned into the value SpreadsheetApp's getValues would give
 * numbers, text and ticks as themselves, errors as their text (e.g. "#N/A"), empty cells as ""
 * pure, no google calls
 *
 * @param cell  the cell, undefined when the row stops before it
 * @returns the value
 */
function _cellValue_(cell: GoogleAppsScript.Sheets.Schema.CellData | undefined): unknown {
  const value = cell?.effectiveValue;
  if (!value) return "";
  if (value.numberValue !== undefined) return value.numberValue;
  if (value.stringValue !== undefined) return value.stringValue;
  if (value.boolValue !== undefined) return value.boolValue;
  return cell?.formattedValue ?? "";
}

/**
 * every tab of a results file with all its values, from a Sheets API response
 * pure, no google calls
 *
 * the API leaves out empty cells at the end of each row and empty rows at the end of the tab,
 * rows are padded so every row is as wide as the widest, like getDataRange().getValues()
 *
 * @param book  the spreadsheet, as returned with its grid data
 * @returns the tabs in file order
 */
function _tabsFromSpreadsheet_(book: GoogleAppsScript.Sheets.Schema.Spreadsheet): ResultsTab[] {
  return (book.sheets ?? []).map((sheet) => {
    const rows = (sheet.data ?? []).flatMap((grid) => grid.rowData ?? []).map((row) =>
      (row.values ?? []).map(_cellValue_)
    );
    const width = Math.max(0, ...rows.map((row) => row.length));
    return {
      title: sheet.properties?.title ?? "",
      values: rows.map((row) => [...row, ...Array(width - row.length).fill("")]),
    };
  });
}

/**
 * fetch many Form Responses files at once: every modified time together, then every file's contents together
 * much faster than reading the files one after another, where each read is mostly waiting
 * modified times are read before the contents, so an edit made during the read is picked up next time
 *
 * @param urls  links to the Form Responses spreadsheets
 * @returns link → its modified time and tabs, or why it could not be fetched
 */
function _readResultsFiles_(urls: string[]): Map<string, FileRead> {
  if (urls.length === 0) return new Map();
  const versions = _readFileVersions_(urls);
  const readable = urls.filter((url) => versions.get(url)?.ok);
  const books = _fetchGoogleJson_(
    readable.map((url) =>
      `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(_spreadsheetIdFromUrl_(url))}`
      + `?includeGridData=true&fields=${encodeURIComponent(RESULTS_FILE_FIELDS)}`
    ),
  );
  const bookByUrl = new Map(readable.map((url, i) => [url, books[i]]));
  return new Map(urls.map((url): [string, FileRead] => {
    const version = versions.get(url);
    if (!version || !version.ok) return [url, { ok: false, reason: version ? version.reason : "not checked" }];
    const book = bookByUrl.get(url);
    if (!book || book.code !== 200) {
      return [url, { ok: false, reason: book ? `could not open file (${_apiError_(book)})` : "not read" }];
    }
    const tabs = _tabsFromSpreadsheet_(book.body as GoogleAppsScript.Sheets.Schema.Spreadsheet);
    return [url, { ok: true, version: version.version, tabs }];
  }));
}

/**
 * file statuses that Refresh Results and Issues imports
 */
const IMPORT_STATUSES: readonly string[] = Object.freeze([
  SOURCE_STATUS.NEW,
  SOURCE_STATUS.EDITED,
  SOURCE_STATUS.REFRESH,
]);

/**
 * outcome of importing one schedule: its games and modified time, or why it could not be read
 */
type ScheduleImport = { ok: true; games: GameRecord[]; version: Date } | { ok: false; reason: string };

/**
 * replace the part of a Message about one file
 * pure, no google calls
 *
 * @param message  the Message cell, parts separated by "; "
 * @param source   which file, e.g. "Schedule"
 * @param reason   what to say about it, "" to say nothing
 * @returns the new Message
 */
function _sourceMessage_(message: string, source: string, reason: string): string {
  const parts = message.split("; ").filter((part) => part !== "" && !part.startsWith(`${source}: `));
  if (reason !== "") parts.push(`${source}: ${reason}`);
  return parts.join("; ");
}

/**
 * work out which files of one kind to import and which to drop
 * pure, no google calls
 *
 * a file is imported once, however many tournaments on the Tournaments tab share it,
 * when any tournament with Include Results ticked has it NEW, EDITED or REFRESH
 * a file no included tournament uses is dropped, with everything imported from it
 *
 * @param events  every event
 * @param url     which file: the schedule or the Form Responses
 * @param status  that file's status
 * @returns links to import, and links still in use by an included tournament
 */
function _filesToImport_(
  events: EventRecord[],
  url: (e: EventRecord) => string,
  status: (e: EventRecord) => string,
): { toImport: string[]; inUse: Set<string> } {
  const included = events.filter((e) => e.includeResults && status(e) !== SOURCE_STATUS.MISSING);
  const toImport = included.filter((e) => IMPORT_STATUSES.includes(status(e))).map(url);
  return { toImport: [...new Set(toImport)], inUse: new Set(included.map(url)) };
}

/**
 * outcome of importing one Form Responses file
 */
type SpiritFormImport =
  | {
    ok: true;
    responses: ResponseRecord[];
    deleted: DeletedFormRecord[];
    problems: RowProblem[];
    version: Date;
  }
  | { ok: false; reason: string };

/**
 * turn a fetched Form Responses file into responses and deleted forms
 * pure, no google calls, the file is fetched beforehand by _readResultsFiles_
 *
 * @param url   link to the Form Responses, stored on every row
 * @param read  the fetched file, undefined if it was not fetched
 * @returns the rows read, or why the file could not be read
 */
function _importSpiritForm_(url: string, read: FileRead | undefined): SpiritFormImport {
  if (!read || !read.ok) return { ok: false, reason: read ? read.reason : "not read" };
  const found = _findFormTabs_(read.tabs);
  if (!found.ok) return { ok: false, reason: found.reason };

  const { responses, problems } = _parseResponseRows_(found.responses.tab.values, found.responses.columns);
  const deleted = found.deleted
    ? _parseDeletedRows_(found.deleted.tab.values, found.deleted.columns, found.deleted.reasonColumn)
    : [];
  return {
    ok: true,
    responses: responses.map((r) => ({
      spiritFormUrl: url,
      sourceRow: r.sourceRow,
      timestamp: r.timestamp,
      scorer: r.scorer,
      receiver: r.receiver,
      ...r.scores,
      comment: r.comment,
      game: "",
      counted: true,
      match: MATCH_RESULTS.NO_SCHEDULE,
    })),
    deleted: deleted.map((f) => _deletedFormRecord_(url, f)),
    problems,
    version: read.version,
  };
}

/**
 * one file's new status after an import, or after Include Results was unticked
 * pure, no google calls
 *
 * @param event       the tournament
 * @param source      "Schedule" or "Spirit Form", for the Message
 * @param importDate  the file's import date now
 * @param result      the file's import, undefined if it was not imported
 * @param note        what to say about a successful import, e.g. skipped rows, "" for nothing
 * @returns the new import date, status and message, null if nothing changes
 */
function _importedSource_(
  event: EventRecord,
  source: string,
  importDate: Date | null,
  result: { ok: true; version: Date } | { ok: false; reason: string } | undefined,
  note: string,
): { importDate: Date | null; status: string; message: string } | null {
  if (!event.includeResults) {
    return importDate === null
      ? null
      : { importDate: null, status: SOURCE_STATUS.NEW, message: _sourceMessage_(event.message, source, "") };
  }
  if (!result) return null;
  if (!result.ok) {
    return { importDate, status: SOURCE_STATUS.ERROR, message: _sourceMessage_(event.message, source, result.reason) };
  }
  return {
    importDate: result.version,
    status: SOURCE_STATUS.OK,
    message: _sourceMessage_(event.message, source, note),
  };
}

/**
 * the Message note for rows of a responses tab that could not be read
 *
 * @param problems  the rows and why
 * @returns e.g. "2 row(s) skipped: row 4 (missing Your Team Name); ...", "" if none
 */
function _problemsNote_(problems: RowProblem[]): string {
  if (problems.length === 0) return "";
  return `${problems.length} row(s) skipped: ${problems.map((p) => `row ${p.sourceRow} (${p.reason})`).join("; ")}`;
}

/**
 * import every included tournament's files marked NEW, EDITED or REFRESH on the Tournaments tab
 * schedules go onto the Games tab, Form Responses onto the Results and Deleted Forms tabs
 * works from the Tournaments tab only, run Refresh Tournaments first to find new or edited files
 *
 * each file is read once, however many tournaments share it
 * a successful import replaces that file's rows and marks it OK on every tournament using it
 * a failed import marks it ERROR and keeps the previous rows
 * tournaments with Include Results unticked lose their rows and go back to NEW
 * every response and deleted form is then matched to its game again, see _matchResponses_
 * the tabs are written before Tournaments, so an interrupted run is simply repeated
 *
 * @returns a one-line summary of the refresh
 */
function _refreshResults_(): string {
  const eventsSheet = _getEventsSheet_();
  const gamesSheet = _getGamesSheet_();
  const responsesSheet = _getResponsesSheet_();
  const deletedSheet = _getDeletedFormsSheet_();
  const events = _readEvents_(eventsSheet);
  if (events.length === 0) return "No tournaments found: run Refresh Tournaments first";

  const schedules = _filesToImport_(events, (e) => e.scheduleUrl, (e) => e.scheduleStatus);
  const forms = _filesToImport_(events, (e) => e.spiritFormUrl, (e) => e.spiritFormStatus);

  const versions = _timed_("read schedule versions", () => _readFileVersions_(schedules.toImport));
  const gamelists = _timed_("read gamelists", () => _readGamelists_(schedules.toImport));
  const scheduleImports: Map<string, ScheduleImport> = new Map(
    schedules.toImport.map((url): [string, ScheduleImport] => {
      const version = versions.get(url);
      const read = gamelists.get(url);
      if (!version || !version.ok) return [url, { ok: false, reason: version ? version.reason : "not checked" }];
      if (!read || !read.ok) return [url, { ok: false, reason: read ? read.reason : "not read" }];
      return [url, { ok: true, games: _parseGamelist_(url, read.values), version: version.version }];
    }),
  );

  const reads = _timed_("read form responses", () => _readResultsFiles_(forms.toImport));
  const formImports: Map<string, SpiritFormImport> = new Map(
    forms.toImport.map((url): [string, SpiritFormImport] => [url, _importSpiritForm_(url, reads.get(url))]),
  );

  const updatedEvents = events.map((event): EventRecord => {
    let updated = event;
    const schedule = _importedSource_(
      updated,
      "Schedule",
      updated.scheduleImportDate,
      scheduleImports.get(updated.scheduleUrl),
      "",
    );
    if (schedule) {
      updated = {
        ...updated,
        scheduleImportDate: schedule.importDate,
        scheduleStatus: schedule.status,
        message: schedule.message,
      };
    }
    const formImport = formImports.get(updated.spiritFormUrl);
    const form = _importedSource_(
      updated,
      "Spirit Form",
      updated.spiritFormImportDate,
      formImport,
      formImport && formImport.ok ? _problemsNote_(formImport.problems) : "",
    );
    if (form) {
      updated = {
        ...updated,
        spiritFormImportDate: form.importDate,
        spiritFormStatus: form.status,
        message: form.message,
      };
    }
    return { ...updated, eventStatus: _eventStatus_(updated) };
  });

  const replacedSchedules = new Set([...scheduleImports].filter(([, r]) => r.ok).map(([url]) => url));
  const newGames = [...scheduleImports.values()].flatMap((r) => (r.ok ? r.games : []));
  const games = _sortGames_([
    ..._readGames_(gamesSheet).filter((g) =>
      schedules.inUse.has(g.scheduleUrl) && !replacedSchedules.has(g.scheduleUrl)
    ),
    ...newGames,
  ], updatedEvents);

  const replacedForms = new Set([...formImports].filter(([, r]) => r.ok).map(([url]) => url));
  const keep = (url: string) => forms.inUse.has(url) && !replacedForms.has(url);
  const newResponses = [...formImports.values()].flatMap((r) => (r.ok ? r.responses : []));
  const responses = _sortResponses_([
    ..._readResponses_(responsesSheet).filter((r) => keep(r.spiritFormUrl)),
    ...newResponses,
  ], updatedEvents);
  const deleted = _sortDeletedForms_([
    ..._readDeletedForms_(deletedSheet).filter((f) => keep(f.spiritFormUrl)),
    ...[...formImports.values()].flatMap((r) => (r.ok ? r.deleted : [])),
  ], updatedEvents);

  const matched = _timed_("match games", () => _matchResponses_(responses, deleted, games, updatedEvents));

  _timed_("write games", () => _writeGames_(gamesSheet, matched.games));
  _timed_("write results", () => _writeResponses_(responsesSheet, matched.responses));
  _timed_("write deleted forms", () => _writeDeletedForms_(deletedSheet, matched.deleted));
  _timed_("write tournaments", () => _writeEvents_(eventsSheet, updatedEvents));

  _rebuildTeamsAndClubs_(matched.responses, _sortEvents_(updatedEvents));

  const summary = (imports: Map<string, { ok: boolean }>, what: string, rows: string) => {
    if (imports.size === 0) return `No ${what} to import`;
    const failed = [...imports.values()].filter((r) => !r.ok).length;
    return `Imported ${imports.size - failed} ${what}, ${rows}`
      + (failed > 0 ? ` (${failed} failed, see Tournaments)` : "");
  };
  return [
    summary(scheduleImports, "schedule(s)", `${newGames.length} game(s)`),
    summary(formImports, "Form Responses", `${newResponses.length} response(s)`),
  ].join("; ");
}
