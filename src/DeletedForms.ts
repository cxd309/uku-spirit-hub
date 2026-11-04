/**
 * name of the Deleted Forms tab
 */
const DELETED_FORMS_SHEET = "Deleted Forms";

/**
 * text for the info row above the Deleted Forms table
 * what the tab is, what to edit, how it refreshes
 */
const DELETED_FORMS_INFO = "Every row on the Deleted forms tab of each included tournament's Form Responses\n\n"
  + "DO NOT EDIT, this is all regenerated on refresh. Filters are kept through a refresh, sorting is put back to the hub's order\n\n"
  + "A record of what the chair moved out of the responses, with their reason. "
  + "Some are scores taken out, others are the original of a response that was edited and is still on Results\n"
  + "Never used in statistics\n\n"
  + "To Refresh run \"Refresh Results and Issues\"";

/**
 * deleted forms tab columns, in order
 * keys are the names used in code
 * values are the header text
 * columns listed in DELETED_FORM_FORMULAS are formulas, the rest hold DeletedFormRecord values
 */
const DELETED_FORM_HEADERS = Object.freeze(
  {
    spiritFormUrl: "Spirit Form URL",
    date: "Date",
    tournament: "Tournament",
    sourceRow: "Source Row",
    timestamp: "Timestamp",
    scorer: "Scoring Team",
    receiver: "Receiving Team",
    rules: "Rules",
    fouls: "Fouls",
    fairMindedness: "Fair-Mindedness",
    attitude: "Attitude",
    communication: "Communication",
    total: "Total",
    comment: "Comment",
    reason: "Reason",
  } as const,
);

/**
 * keys of DeletedFormRecord in column order
 */
const DELETED_FORM_KEYS = Object.keys(DELETED_FORM_HEADERS) as (keyof typeof DELETED_FORM_HEADERS)[];

/**
 * one row of the Deleted Forms tab
 */
interface DeletedFormRecord {
  /** link to the tournament's Form Responses (links to the Events tab) */
  spiritFormUrl: string;
  /** row on the Deleted forms tab */
  sourceRow: number;
  /** when the form was submitted, null if blank */
  timestamp: Date | null;
  /** team giving the score */
  scorer: string;
  /** team receiving the score */
  receiver: string;
  /** rules Knowledge and Use, null if missing */
  rules: number | null;
  /** fouls and Body Contact, null if missing */
  fouls: number | null;
  /** fair-Mindedness, null if missing */
  fairMindedness: number | null;
  /** positive Attitude and Self-Control, null if missing */
  attitude: number | null;
  /** communication, null if missing */
  communication: number | null;
  /** sum of the five scores, null if any is missing */
  total: number | null;
  /** comment, "" if none */
  comment: string;
  /** the chair's reason for deleting it */
  reason: string;
}

/**
 * formula columns of the Deleted Forms tab, looking the tournament up on the Tournaments tab
 * each row gets its own formula, so sorting the tab never breaks them
 */
const DELETED_FORM_FORMULAS: Readonly<Partial<Record<keyof typeof DELETED_FORM_HEADERS, (row: number) => string>>> =
  Object.freeze({
    date: (row: number) => _deletedFormLookup_(row, "startDate", "\"\""),
    tournament: (row: number) => _deletedFormLookup_(row, "eventName", "\"(unknown event)\""),
  });

/**
 * a formula looking one Deleted Forms row's tournament up on the Tournaments tab
 *
 * @param row        1-based sheet row the formula is for
 * @param key        Tournaments column to return
 * @param otherwise  formula text when the tournament is not found
 * @returns the formula
 */
function _deletedFormLookup_(row: number, key: keyof typeof EVENT_HEADERS, otherwise: string): string {
  const url = `$${_columnLetter_(DELETED_FORM_KEYS.indexOf("spiritFormUrl") + 1)}${row}`;
  const events = (k: keyof typeof EVENT_HEADERS) => _columnBelowHeader_(EVENTS_SHEET, EVENT_KEYS.indexOf(k) + 1);
  return `=IFERROR(XLOOKUP(${url}, ${events("spiritFormUrl")}, ${events(key)}), ${otherwise})`;
}

/**
 * get the Deleted Forms tab, creating it with header row on first use
 *
 * @returns the Deleted Forms tab
 */
function _getDeletedFormsSheet_(): GoogleAppsScript.Spreadsheet.Sheet {
  const headers: readonly string[] = Object.values(DELETED_FORM_HEADERS);
  return _getOrCreateSheet_(SpreadsheetApp.getActiveSpreadsheet(), DELETED_FORMS_SHEET, headers, DELETED_FORMS_INFO);
}

/**
 * the record for one deleted form of a tournament
 * pure, no google calls
 *
 * @param spiritFormUrl  the tournament's Form Responses link
 * @param form           the row as read from the Deleted forms tab
 * @returns the record
 */
function _deletedFormRecord_(spiritFormUrl: string, form: DeletedForm): DeletedFormRecord {
  const scores = SCORE_KEYS.map((key) => form.scores[key]);
  return {
    spiritFormUrl,
    sourceRow: form.sourceRow,
    timestamp: form.timestamp,
    scorer: form.scorer,
    receiver: form.receiver,
    ...form.scores,
    total: scores.every((s) => s !== null) ? scores.reduce((a: number, b) => a + (b ?? 0), 0) : null,
    comment: form.comment,
    reason: form.reason,
  };
}

/**
 * @param row  values of one data row, in DELETED_FORM_KEYS order
 * @returns the record
 */
function _deletedFormFromRow_(row: unknown[]): DeletedFormRecord {
  const cell = (key: keyof typeof DELETED_FORM_HEADERS) => row[DELETED_FORM_KEYS.indexOf(key)];
  const score = (
    key: keyof typeof DELETED_FORM_HEADERS,
  ) => (typeof cell(key) === "number" ? cell(key) as number : null);
  return {
    spiritFormUrl: String(cell("spiritFormUrl")),
    sourceRow: Number(cell("sourceRow")),
    timestamp: cell("timestamp") instanceof Date ? cell("timestamp") as Date : null,
    scorer: String(cell("scorer")),
    receiver: String(cell("receiver")),
    rules: score("rules"),
    fouls: score("fouls"),
    fairMindedness: score("fairMindedness"),
    attitude: score("attitude"),
    communication: score("communication"),
    total: score("total"),
    comment: String(cell("comment")),
    reason: String(cell("reason")),
  };
}

/**
 * @param form  the record
 * @returns values in DELETED_FORM_KEYS order, "" in formula columns (filled by _fillFormulaColumns_) and for blanks
 */
function _deletedFormToRow_(form: DeletedFormRecord): unknown[] {
  return DELETED_FORM_KEYS.map((key) => {
    if (DELETED_FORM_FORMULAS[key]) return "";
    const value = form[key as keyof DeletedFormRecord];
    return value === null ? "" : value;
  });
}

/**
 * @param sheet  the Deleted Forms tab
 * @returns every deleted form on the tab
 */
function _readDeletedForms_(sheet: GoogleAppsScript.Spreadsheet.Sheet): DeletedFormRecord[] {
  const rowCount = sheet.getLastRow() - HEADER_ROW;
  if (rowCount < 1) return [];
  return sheet
    .getRange(DATA_ROW, 1, rowCount, DELETED_FORM_KEYS.length)
    .getValues()
    .map(_deletedFormFromRow_)
    .filter((f) => f.spiritFormUrl !== "");
}

/**
 * replace the contents of the Deleted Forms tab, resizing it to fit
 *
 * @param sheet  the Deleted Forms tab
 * @param forms  all deleted forms, in display order
 */
function _writeDeletedForms_(sheet: GoogleAppsScript.Spreadsheet.Sheet, forms: DeletedFormRecord[]) {
  _writeTable_(sheet, forms.map(_deletedFormToRow_), DELETED_FORM_KEYS.length);
  if (forms.length === 0) return;
  _fillFormulaColumns_(sheet, DELETED_FORM_KEYS, DELETED_FORM_FORMULAS, forms.length);
  const range = (key: keyof typeof DELETED_FORM_HEADERS) =>
    sheet.getRange(DATA_ROW, DELETED_FORM_KEYS.indexOf(key) + 1, forms.length, 1);
  range("date").setNumberFormat("yyyy-mm-dd");
  range("timestamp").setNumberFormat("yyyy-mm-dd hh:mm:ss");
}

/**
 * deleted forms in the date order of their events, oldest first, then as they are on each Deleted forms tab
 * pure, returns a new array
 *
 * @param forms   deleted forms in any order
 * @param events  events, used to find each form's date and tournament
 * @returns deleted forms in date order
 */
function _sortDeletedForms_(forms: DeletedFormRecord[], events: EventRecord[]): DeletedFormRecord[] {
  const eventByUrl = new Map(events.map((e) => [e.spiritFormUrl, e]));
  return [...forms].sort((a, b) => {
    const ea = eventByUrl.get(a.spiritFormUrl);
    const eb = eventByUrl.get(b.spiritFormUrl);
    return _compareDates_(ea?.startDate ?? null, eb?.startDate ?? null)
      || (ea?.eventName ?? "").localeCompare(eb?.eventName ?? "")
      || a.spiritFormUrl.localeCompare(b.spiritFormUrl)
      || a.sourceRow - b.sourceRow;
  });
}
