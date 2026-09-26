/**
 * Outcome of importing one results file.
 */
type ImportResult = {
  ok: true;
  responses: ResponseRecord[];
  problems: RowProblem[];
  version: Date;
} | {
  ok: false;
  reason: string;
};

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
 * fetch many results files at once: every modified time together, then every file's contents together
 * much faster than reading the files one after another, where each read is mostly waiting
 * modified times are read before the contents, so an edit made during the read is picked up next time
 *
 * @param fileIds  Drive file IDs of the results files
 * @returns file id → its modified time and tabs, or why it could not be fetched
 */
function _readResultsFiles_(fileIds: string[]): Map<string, FileRead> {
  if (fileIds.length === 0) return new Map();
  const modified = _fetchGoogleJson_(
    fileIds.map((id) =>
      `https://www.googleapis.com/drive/v3/files/${encodeURIComponent(id)}?fields=modifiedTime&supportsAllDrives=true`
    ),
  );
  const books = _fetchGoogleJson_(
    fileIds.map((id) =>
      `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(id)}`
      + `?includeGridData=true&fields=${encodeURIComponent(RESULTS_FILE_FIELDS)}`
    ),
  );
  return new Map(fileIds.map((id, i): [string, FileRead] => {
    const failed = [modified[i], books[i]].find((r) => r.code !== 200);
    if (failed) return [id, { ok: false, reason: _apiError_(failed) }];
    const file = modified[i].body as GoogleAppsScript.Drive_v3.Drive.V3.Schema.File;
    const book = books[i].body as GoogleAppsScript.Sheets.Schema.Spreadsheet;
    return [id, { ok: true, version: new Date(file.modifiedTime ?? 0), tabs: _tabsFromSpreadsheet_(book) }];
  }));
}

/**
 * why an event's results file must not be imported, from its folder and file names
 * pure, no google calls
 *
 * @param event  the event
 * @returns the reason, "" if it can be imported
 */
function _importProblem_(event: EventRecord): string {
  if (event.date === null) {
    return "folder name must be \"YYYYMMDD Tournament name\", fix it then run Refresh Tournaments";
  }
  if (!event.fileName.toLowerCase().includes(RESULTS_FILE_TEXT.toLowerCase())) {
    return `file name "${event.fileName}" does not contain "${RESULTS_FILE_TEXT}", fix it then run Refresh Tournaments`;
  }
  return "";
}

/**
 * turn an event's fetched results file into response records
 * pure, no google calls, the file is fetched beforehand by _readResultsFiles_
 *
 * a file that could not be fetched (deleted, no permission) is returned as a reason,
 * so one bad file cannot stop the whole refresh
 *
 * @param event  the event to import
 * @param read   its fetched file, undefined if it was not fetched
 * @returns the responses read, or why the file could not be read
 */
function _importFile_(event: EventRecord, read: FileRead | undefined): ImportResult {
  const problem = _importProblem_(event);
  if (problem !== "") return { ok: false, reason: problem };
  if (!read || !read.ok) {
    const message = read ? read.reason : "not fetched";
    return { ok: false, reason: `could not open file (${message}), try Refresh Tournaments first` };
  }

  const found = _findBreakdownTab_(read.tabs);
  if (!found.ok) return { ok: false, reason: found.reason };

  const { responses, problems } = _parseBreakdownRows_(found.tab.values, found.columns);
  return {
    ok: true,
    responses: responses.map((r) => ({
      fileId: event.fileId,
      sourceRow: r.sourceRow,
      scorer: r.scorer,
      receiver: r.receiver,
      ...r.scores,
      comment: r.comment,
    })),
    problems: problems,
    version: read.version,
  };
}

/**
 * import every tournament marked NEW or REFRESH on the Tournaments tab
 * works from the Tournaments tab only, run Refresh Tournaments first to find new or edited files
 *
 * successful imports replace that event's responses and clear its status
 * failed imports mark the event ERROR and keep its previous responses
 * responses are written before Tournaments, so an interrupted run is simply repeated
 *
 * @returns a one-line summary of the refresh
 */
function _refreshResults_(): string {
  const eventsSheet = _getEventsSheet_();
  const responsesSheet = _getResponsesSheet_();
  const events = _readEvents_(eventsSheet);
  if (events.length === 0) return "No tournaments found: run Refresh Tournaments first";

  const toImport = events.filter((e) => e.status === EVENT_STATUS.NEW || e.status === EVENT_STATUS.REFRESH);
  if (toImport.length === 0) return "Nothing to refresh: no tournaments are NEW or REFRESH";

  const readable = toImport.filter((e) => _importProblem_(e) === "");
  const reads = _timed_("read files", () => _readResultsFiles_(readable.map((e) => e.fileId)));
  const results: Map<string, ImportResult> = new Map(
    toImport.map((e) => [e.fileId, _importFile_(e, reads.get(e.fileId))]),
  );

  const updatedEvents = events.map((event) => {
    const result = results.get(event.fileId);
    if (!result) return event;
    if (!result.ok) return { ...event, status: EVENT_STATUS.ERROR, message: result.reason };
    const skipped = result.problems.length === 0
      ? ""
      : `${result.problems.length} row(s) skipped: ${
        result.problems.map((p) => `row ${p.sourceRow} (${p.reason})`).join("; ")
      }`;
    return { ...event, status: EVENT_STATUS.OK, message: skipped, importedVersion: result.version };
  });

  const replacedIds = new Set([...results].filter(([, r]) => r.ok).map(([id]) => id));
  const imported = [...results.values()].flatMap((r) => (r.ok ? r.responses : []));
  const responses = _sortResponses_([
    ..._readResponses_(responsesSheet).filter((r) => !replacedIds.has(r.fileId)),
    ...imported,
  ], updatedEvents);

  _timed_("write results", () => _writeResponses_(responsesSheet, responses));
  _timed_("write tournaments", () => _writeEvents_(eventsSheet, updatedEvents));

  const internationalIds = new Set(updatedEvents.filter((e) => e.international).map((e) => e.fileId));
  const teamsSheet = _getTeamsSheet_();
  const sortedEvents = _sortEvents_(updatedEvents);
  const teamClubs = _timed_("write teams", () =>
    _writeTeams_(
      teamsSheet,
      _mergeTeams_(_readTeams_(teamsSheet), _teamNamesFromResponses_(responses, internationalIds)),
      responses,
      sortedEvents,
    ));
  _timed_("rebuild clubs", () => _rebuildClubs_(teamClubs, responses, sortedEvents));

  const failed = [...results.values()].filter((r) => !r.ok).length;
  return `Imported ${results.size - failed} tournament(s), ${imported.length} response(s)`
    + (failed > 0 ? `; ${failed} failed (see Tournaments tab)` : "");
}
