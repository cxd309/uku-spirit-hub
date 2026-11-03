/**
 * name of the tab in the Schedule Manager that lists every tournament
 */
const SCHEDULE_LIST_SHEET = "Schedule List";

/**
 * header text of the Schedule List columns the hub reads
 * keys are names used in code
 * matched to row 1 ignoring case and surrounding spaces, other columns are ignored
 */
const SCHEDULE_LIST_HEADERS = Object.freeze({
  eventName: "Event Name",
  startDate: "Start date",
  endDate: "End date",
  scheduleUrl: "Schedule URL",
  spiritFormUrl: "SOTG Sheet",
});

/**
 * one tournament from the Schedule List, as the hub needs it
 */
interface ScheduleListRow {
  /** Event Name */
  eventName: string;
  /** Start date */
  startDate: Date;
  /** End date, null if blank */
  endDate: Date | null;
  /** link to the tournament's schedule spreadsheet */
  scheduleUrl: string;
  /** link to the tournament's Form Responses spreadsheet (the key) */
  spiritFormUrl: string;
  /** 1-based row on the Schedule List, for messages */
  sourceRow: number;
}

/**
 * the tournaments read from the Schedule List, and the rows left out because they could not be read
 */
interface ScheduleListRead {
  /** rows with both links and a start date within the date range */
  rows: ScheduleListRow[];
  /** 1-based row numbers with both links but no start date */
  noStartDate: number[];
}

/**
 * read every value on the Schedule List tab of the Schedule Manager
 * uses the Sheets API rather than SpreadsheetApp, which needs permission to edit every spreadsheet just to open one
 * dates come back as serial numbers, _parseScheduleList_ turns them into dates
 *
 * @param url  link to the Schedule Manager, from the Hub settings
 * @returns every value from row 1
 * @throws {Error} if the spreadsheet cannot be opened or has no Schedule List tab
 */
function _readScheduleListValues_(url: string): unknown[][] {
  const id = _spreadsheetIdFromUrl_(url);
  if (id === "") throw new Error(`"${HUB_SETTINGS.scheduleManagerUrl.label}" is not a Google Sheets link: ${url}`);
  const [response] = _fetchGoogleJson_([
    `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(id)}/values/`
    + `${encodeURIComponent(`'${SCHEDULE_LIST_SHEET}'`)}`
    + "?valueRenderOption=UNFORMATTED_VALUE&dateTimeRenderOption=SERIAL_NUMBER",
  ]);
  if (response.code === 400) throw new Error(`The Schedule Manager has no "${SCHEDULE_LIST_SHEET}" tab`);
  if (response.code !== 200) throw new Error(`Could not open the Schedule Manager (${_apiError_(response)})`);
  return (response.body as GoogleAppsScript.Sheets.Schema.ValueRange).values ?? [];
}

/**
 * a Schedule List cell as a date
 * pure, no google calls
 *
 * @param value  a date, or a Sheets serial number (days since 1899-12-30) as the Sheets API returns dates
 * @returns the date at midnight, null if the cell is not a date
 */
function _scheduleDate_(value: unknown): Date | null {
  if (value instanceof Date) return value;
  if (typeof value !== "number" || !Number.isFinite(value)) return null;
  return new Date(1899, 11, 30 + Math.floor(value));
}

/**
 * turn the values of the Schedule List into tournaments
 * pure, no google calls
 *
 * rows without a Schedule URL or SOTG Sheet are ignored, as are rows starting outside the date range
 * rows with both links but no start date are counted, so they can be reported
 *
 * @param values     every value on the Schedule List from row 1, from _readScheduleListValues_
 * @param startDate  first start date to include
 * @param endDate    last start date to include
 * @returns the tournaments in sheet order, and the rows with no start date
 * @throws {Error} if a column the hub needs is missing from row 1
 */
function _parseScheduleList_(values: unknown[][], startDate: Date, endDate: Date): ScheduleListRead {
  const header = (values[0] ?? []).map((cell) => String(cell).trim().toLowerCase());
  const keys = Object.keys(SCHEDULE_LIST_HEADERS) as (keyof typeof SCHEDULE_LIST_HEADERS)[];
  const columns = {} as Record<keyof typeof SCHEDULE_LIST_HEADERS, number>;
  for (const key of keys) {
    const index = header.indexOf(SCHEDULE_LIST_HEADERS[key].toLowerCase());
    if (index < 0) throw new Error(`The "${SCHEDULE_LIST_SHEET}" tab has no "${SCHEDULE_LIST_HEADERS[key]}" column`);
    columns[key] = index;
  }

  const first = _dayNumber_(startDate);
  const last = _dayNumber_(endDate);
  const rows: ScheduleListRow[] = [];
  const noStartDate: number[] = [];
  values.slice(1).forEach((row, i) => {
    const text = (key: keyof typeof SCHEDULE_LIST_HEADERS) => String(row[columns[key]] ?? "").trim();
    const scheduleUrl = text("scheduleUrl");
    const spiritFormUrl = text("spiritFormUrl");
    if (scheduleUrl === "" || spiritFormUrl === "") return;

    const start = _scheduleDate_(row[columns.startDate]);
    if (start === null) {
      noStartDate.push(i + 2);
      return;
    }
    if (_dayNumber_(start) < first || _dayNumber_(start) > last) return;

    rows.push({
      eventName: text("eventName"),
      startDate: start,
      endDate: _scheduleDate_(row[columns.endDate]),
      scheduleUrl,
      spiritFormUrl,
      sourceRow: i + 2,
    });
  });
  return { rows, noStartDate };
}

/**
 * a date's calendar day as a whole number, so dates can be compared ignoring the time of day
 * pure, no google calls
 *
 * @param date  the date
 * @returns days since 1970-01-01 for the date's own year, month and day
 */
function _dayNumber_(date: Date): number {
  return Math.round(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()) / 86400000);
}

/**
 * the spreadsheet ID inside a Google Sheets link
 * pure, no google calls
 *
 * @param url  e.g. https://docs.google.com/spreadsheets/d/<id>/edit?usp=drivesdk
 * @returns the ID, "" if the link is not a Google Sheets link
 */
function _spreadsheetIdFromUrl_(url: string): string {
  const match = /\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/.exec(url);
  return match ? match[1] : "";
}

/**
 * a file's last-modified time, or why it could not be read
 */
type FileVersion = { ok: true; version: Date } | { ok: false; reason: string };

/**
 * fetch the last-modified time of many spreadsheets at once
 * links that are not Google Sheets links are not fetched and fail straight away
 *
 * @param urls  spreadsheet links, duplicates are fetched once
 * @returns link → its modified time, or why it could not be read
 */
function _readFileVersions_(urls: string[]): Map<string, FileVersion> {
  const unique = [...new Set(urls)];
  const versions: Map<string, FileVersion> = new Map();
  const toFetch: string[] = [];
  for (const url of unique) {
    if (_spreadsheetIdFromUrl_(url) === "") versions.set(url, { ok: false, reason: "not a Google Sheets link" });
    else toFetch.push(url);
  }
  if (toFetch.length === 0) return versions;

  const responses = _fetchGoogleJson_(
    toFetch.map((url) =>
      `https://www.googleapis.com/drive/v3/files/${encodeURIComponent(_spreadsheetIdFromUrl_(url))}`
      + "?fields=modifiedTime&supportsAllDrives=true"
    ),
  );
  toFetch.forEach((url, i) => {
    const response = responses[i];
    if (response.code === 404) {
      versions.set(url, { ok: false, reason: "file not found, or not shared with you" });
      return;
    }
    if (response.code !== 200) {
      versions.set(url, { ok: false, reason: `could not open file (${_apiError_(response)})` });
      return;
    }
    const file = response.body as GoogleAppsScript.Drive_v3.Drive.V3.Schema.File;
    versions.set(url, { ok: true, version: new Date(file.modifiedTime ?? 0) });
  });
  return versions;
}
