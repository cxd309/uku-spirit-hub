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
 * null (invalid folder name) sorts last
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
 * write rows below the header, resize the tab to fit, and put a filter on the table
 *
 * @param sheet    tab to write
 * @param rows     data rows, formulas included
 * @param columns  number of columns in the table
 */
function _writeTable_(sheet: GoogleAppsScript.Spreadsheet.Sheet, rows: unknown[][], columns: number) {
  _fitSheet_(sheet, HEADER_ROW + rows.length, columns);
  if (rows.length > 0) sheet.getRange(DATA_ROW, 1, rows.length, columns).setValues(rows);
  _applyFilter_(sheet, rows.length, columns);
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
 * put a fresh filter over the header row and data rows
 * any existing filter is removed first, so people's filter settings are reset on each write
 *
 * @param sheet     tab to filter
 * @param dataRows  number of data rows
 * @param columns   number of columns in the table
 */
function _applyFilter_(sheet: GoogleAppsScript.Spreadsheet.Sheet, dataRows: number, columns: number) {
  sheet.getFilter()?.remove();
  sheet.getRange(HEADER_ROW, 1, 1 + Math.max(dataRows, 1), columns).createFilter();
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
