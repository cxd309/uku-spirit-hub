/**
 * get a tab by name, creating it if it does not exist
 * a new tab gets a wrapped info row, then a bold header row, both frozen
 *
 * if the tab exists, its header row must match `headers` exactly
 * code and sheet cannot silently disagree
 *
 * @param ss       spreadsheet to examine
 * @param name     tab name
 * @param headers  expected header row
 * @param info     text for the info row
 *                 line breaks separate points
 * @returns the tab
 * @throws {Error} if the tab exists but its header row differs from `headers`
 */
function _getOrCreateSheet_(
  ss: GoogleAppsScript.Spreadsheet.Spreadsheet,
  name: string,
  headers: readonly string[],
  info: string,
): GoogleAppsScript.Spreadsheet.Sheet {
  const existing = ss.getSheetByName(name);
  if (existing) {
    const actual = existing.getRange(HEADER_ROW, 1, 1, headers.length).getValues()[0].map(String);
    if (actual.join("|") !== headers.join("|")) {
      throw new Error(
        `Tab "${name}" has unexpected headers: expected [${headers.join(", ")}], found [${actual.join(", ")}]`,
      );
    }
    return existing;
  }

  const sheet = ss.insertSheet(name);
  _fitSheet_(sheet, DATA_ROW, headers.length);
  sheet
    .getRange(INFO_ROW, 1, 1, headers.length)
    .merge()
    .setValue(info)
    .setWrap(true)
    .setVerticalAlignment("top");
  sheet.getRange(HEADER_ROW, 1, 1, headers.length).setValues([[...headers]]).setFontWeight("bold");
  sheet.setFrozenRows(HEADER_ROW);
  return sheet;
}

/**
 * resize a tab to exactly the given number of rows and columns
 * adding or deleting rows and columns at the end as needed
 *
 * always keeps at least one row below any frozen rows, because Sheets does not
 * allow deleting every unfrozen row.
 *
 * @param sheet    tab to resize
 * @param rows     total rows wanted, including the info and header rows
 * @param columns  total columns wanted
 */
function _fitSheet_(sheet: GoogleAppsScript.Spreadsheet.Sheet, rows: number, columns: number) {
  const wantRows = Math.max(rows, sheet.getFrozenRows() + 1);
  const maxRows = sheet.getMaxRows();
  if (maxRows < wantRows) {
    sheet.insertRowsAfter(maxRows, wantRows - maxRows);
  } else if (maxRows > wantRows) {
    sheet.deleteRows(wantRows + 1, maxRows - wantRows);
  }

  const maxColumns = sheet.getMaxColumns();
  if (maxColumns < columns) {
    sheet.insertColumnsAfter(maxColumns, columns - maxColumns);
  } else if (maxColumns > columns) {
    sheet.deleteColumns(columns + 1, maxColumns - columns);
  }
}

/**
 * convert a 1-based column number to its letter(s), e.g. 1 → "A", 28 → "AB"
 *
 * @param column  1-based column number
 * @returns column letters
 */
function _columnLetter_(column: number): string {
  let letters = "";
  let n = column;
  while (n > 0) {
    const remainder = (n - 1) % 26;
    letters = String.fromCharCode(65 + remainder) + letters;
    n = Math.floor((n - 1) / 26);
  }
  return letters;
}

/**
 * absolute reference to a column of another tab, below its header row
 * e.g. 'Teams'!$A$3:$A
 * use when the header text must not be counted as data
 *
 * @param sheetName  tab name
 * @param column     1-based column number
 * @returns the reference, for use inside formulas
 */
function _columnBelowHeader_(sheetName: string, column: number): string {
  const letter = _columnLetter_(column);
  return `'${sheetName}'!$${letter}$${DATA_ROW}:$${letter}`;
}

/**
 * compare two dates for sorting, oldest first
 * null (no date) sorts last
 *
 * @param a  first date
 * @param b  second date
 * @returns negative if a comes first, positive if b comes first, 0 if equal
 */
function _compareDates_(a: Date | null, b: Date | null): number {
  if (a === null || b === null) return (a === null ? 1 : 0) - (b === null ? 1 : 0);
  return a.getTime() - b.getTime();
}

/**
 * background of columns the script fills, light grey 1 as on the Schedule Manager
 * columns a person edits are left white
 */
const GENERATED_BACKGROUND = "#d9d9d9";

/**
 * whether a column holds links, which are clipped rather than wrapped so rows stay short
 *
 * @param header  the column's header text
 * @returns true when the header ends in URL
 */
function _isUrlColumn_(header: unknown): boolean {
  return /URL$/.test(String(header).trim());
}

/**
 * style rows of a table: generated columns grey and columns a person edits white,
 * link columns clipped and every other column wrapped
 * link columns are found from the header row, one call each for the backgrounds and the wrapping
 *
 * @param sheet     tab to style
 * @param firstRow  first row to style, e.g. the header row
 * @param rowCount  how many rows
 * @param columns   number of columns in the table
 * @param editable  0-based columns a person edits
 */
function _styleRows_(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  firstRow: number,
  rowCount: number,
  columns: number,
  editable: readonly number[],
) {
  if (rowCount < 1) return;
  const range = sheet.getRange(firstRow, 1, rowCount, columns);
  const shading = Array.from({ length: columns }, (_, i) => (editable.includes(i) ? null : GENERATED_BACKGROUND));
  range.setBackgrounds(Array.from({ length: rowCount }, () => shading));
  const headers = sheet.getRange(HEADER_ROW, 1, 1, columns).getValues()[0];
  const wrapping = headers.map((header) =>
    _isUrlColumn_(header) ? SpreadsheetApp.WrapStrategy.CLIP : SpreadsheetApp.WrapStrategy.WRAP
  );
  range.setWrapStrategies(Array.from({ length: rowCount }, () => wrapping));
}

/**
 * write rows below the header, resize the tab to fit, put a filter on the table,
 * and style the header and data rows, see _styleRows_
 *
 * @param sheet     tab to write
 * @param rows      data rows, formulas included
 * @param columns   number of columns in the table
 * @param editable  0-based columns a person edits, left white, every other column is shaded
 */
function _writeTable_(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  rows: unknown[][],
  columns: number,
  editable: readonly number[] = [],
) {
  _fitSheet_(sheet, HEADER_ROW + rows.length, columns);
  if (rows.length > 0) sheet.getRange(DATA_ROW, 1, rows.length, columns).setValues(rows);
  _applyFilter_(sheet, rows.length, columns);
  _styleRows_(sheet, HEADER_ROW, 1 + rows.length, columns, editable);
}

/**
 * put a formula in every data row of each formula column, one call per column
 * each formula is built for the first data row and Sheets copies it down, adjusting relative row references
 * like a fill-down, so every row still has its own formula (safe to sort) without sending one formula per cell
 *
 * @param sheet     tab to fill, already written by _writeTable_
 * @param keys      the table's column keys, in column order
 * @param formulas  formula builder for each formula column, given a 1-based sheet row
 * @param rows      number of data rows
 */
function _fillFormulaColumns_<K extends string>(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  keys: readonly K[],
  formulas: Readonly<Partial<Record<K, (row: number) => string>>>,
  rows: number,
) {
  if (rows === 0) return;
  keys.forEach((key, i) => {
    const formula = formulas[key];
    if (formula) sheet.getRange(DATA_ROW, i + 1, rows, 1).setFormula(formula(DATA_ROW));
  });
}

/**
 * put a fresh filter over the header row and data rows, keeping what people had filtered
 * a filter cannot grow with the table, so the old one is removed and a new one made,
 * each column's criteria are carried over to the column with the same header text
 *
 * @param sheet     tab to filter
 * @param dataRows  number of data rows
 * @param columns   number of columns in the table
 */
function _applyFilter_(sheet: GoogleAppsScript.Spreadsheet.Sheet, dataRows: number, columns: number) {
  const headers = sheet.getRange(HEADER_ROW, 1, 1, columns).getValues()[0].map((h) => String(h).trim());
  const kept: Map<string, GoogleAppsScript.Spreadsheet.FilterCriteria> = new Map();
  const old = sheet.getFilter();
  if (old) {
    const range = old.getRange();
    for (let column = range.getColumn(); column <= range.getLastColumn(); column++) {
      const criteria = old.getColumnFilterCriteria(column);
      const header = headers[column - 1];
      if (criteria && header) kept.set(header, criteria.copy().build());
    }
    old.remove();
  }
  const filter = sheet.getRange(HEADER_ROW, 1, 1 + Math.max(dataRows, 1), columns).createFilter();
  headers.forEach((header, i) => {
    const criteria = kept.get(header);
    if (criteria) filter.setColumnFilterCriteria(i + 1, criteria);
  });
}

/**
 * date as yyyy-mm-dd, using the date's own calendar day
 *
 * @param date  the date
 * @returns e.g. "2025-11-01"
 */
function _isoDate_(date: Date): string {
  const pad = (n: number) => String(n).padStart(2, "0");
  return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}`;
}

/**
 * date and time as yyyy-mm-dd hh:mm, using the date's own calendar day and clock
 *
 * @param date  the date
 * @returns e.g. "2025-11-01 14:05"
 */
function _isoDateTime_(date: Date): string {
  const pad = (n: number) => String(n).padStart(2, "0");
  return `${_isoDate_(date)} ${pad(date.getHours())}:${pad(date.getMinutes())}`;
}
