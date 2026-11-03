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
 * import the results of every included tournament
 * being rebuilt to read the Form Responses listed on the Tournaments tab
 *
 * @returns a one-line summary of the refresh
 */
function _refreshResults_(): string {
  return "Refresh Results is being rebuilt to read the Form Responses, it is not available in this version";
}
