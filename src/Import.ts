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
 * file statuses that Refresh Results imports
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
 * work out which schedules to import and which to drop
 * pure, no google calls
 *
 * a schedule is imported once, however many tournaments on the Tournaments tab share it,
 * when any tournament with Include Results ticked has it NEW, EDITED or REFRESH
 * a schedule no included tournament uses is dropped, with its games
 *
 * @param events  every event
 * @returns schedule links to import, and links still in use by an included tournament
 */
function _schedulesToImport_(events: EventRecord[]): { toImport: string[]; inUse: Set<string> } {
  const included = events.filter((e) => e.includeResults && e.scheduleStatus !== SOURCE_STATUS.MISSING);
  const toImport = included.filter((e) => IMPORT_STATUSES.includes(e.scheduleStatus)).map((e) => e.scheduleUrl);
  return { toImport: [...new Set(toImport)], inUse: new Set(included.map((e) => e.scheduleUrl)) };
}

/**
 * import every included tournament's files marked NEW, EDITED or REFRESH on the Tournaments tab
 * so far only schedules are imported, onto the Games tab, Form Responses are not read yet
 * works from the Tournaments tab only, run Refresh Tournaments first to find new or edited files
 *
 * a successful import replaces that schedule's games and marks every tournament using it OK
 * a failed import marks them ERROR and keeps the previous games
 * tournaments with Include Results unticked lose their games and go back to NEW
 * games are written before Tournaments, so an interrupted run is simply repeated
 *
 * @returns a one-line summary of the refresh
 */
function _refreshResults_(): string {
  const eventsSheet = _getEventsSheet_();
  const gamesSheet = _getGamesSheet_();
  const events = _readEvents_(eventsSheet);
  if (events.length === 0) return "No tournaments found: run Refresh Tournaments first";

  const { toImport, inUse } = _schedulesToImport_(events);
  const versions = _timed_("read file versions", () => _readFileVersions_(toImport));
  const gamelists = _timed_("read gamelists", () => _readGamelists_(toImport));
  const imported: Map<string, ScheduleImport> = new Map(toImport.map((url): [string, ScheduleImport] => {
    const version = versions.get(url);
    const read = gamelists.get(url);
    if (!version || !version.ok) return [url, { ok: false, reason: version ? version.reason : "not checked" }];
    if (!read || !read.ok) return [url, { ok: false, reason: read ? read.reason : "not read" }];
    return [url, { ok: true, games: _parseGamelist_(url, read.values), version: version.version }];
  }));

  const updatedEvents = events.map((event): EventRecord => {
    let updated = event;
    const result = imported.get(event.scheduleUrl);
    if (!event.includeResults && event.scheduleImportDate !== null) {
      updated = { ...event, scheduleImportDate: null, scheduleStatus: SOURCE_STATUS.NEW };
    } else if (event.includeResults && result) {
      updated = result.ok
        ? {
          ...event,
          scheduleImportDate: result.version,
          scheduleStatus: SOURCE_STATUS.OK,
          message: _sourceMessage_(event.message, "Schedule", ""),
        }
        : {
          ...event,
          scheduleStatus: SOURCE_STATUS.ERROR,
          message: _sourceMessage_(event.message, "Schedule", result.reason),
        };
    }
    return { ...updated, eventStatus: _eventStatus_(updated) };
  });

  const replaced = new Set([...imported].filter(([, r]) => r.ok).map(([url]) => url));
  const newGames = [...imported.values()].flatMap((r) => (r.ok ? r.games : []));
  const games = _sortGames_([
    ..._readGames_(gamesSheet).filter((g) => inUse.has(g.scheduleUrl) && !replaced.has(g.scheduleUrl)),
    ...newGames,
  ], updatedEvents);

  _timed_("write games", () => _writeGames_(gamesSheet, games));
  _timed_("write tournaments", () => _writeEvents_(eventsSheet, updatedEvents));

  const failed = [...imported.values()].filter((r) => !r.ok).length;
  const schedules = toImport.length === 0
    ? "No schedules to import"
    : `Imported ${toImport.length - failed} schedule(s), ${newGames.length} game(s)`
      + (failed > 0 ? `; ${failed} failed (see Tournaments tab)` : "");
  return `${schedules}; Form Responses are not imported yet`;
}
