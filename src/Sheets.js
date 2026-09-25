/**
 * get a tab by name, creating it if it does not exist
 * a new tab gets a wrapped info row, then a bold header row, both frozen
 *
 * if the tab exists, its header row must match `headers` exactly
 * code and sheet cannot silently disagree
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss       spreadsheet to examine
 * @param {string}                                   name     tab name
 * @param {readonly string[]}                        headers  expected header row
 * @param {string}                                   info     text for the info row
 *                                                            line breaks separate points
 * @returns {{sheet: GoogleAppsScript.Spreadsheet.Sheet, created: boolean}} the tab and if created
 * @throws {Error} if the tab exists but its header row differs from `headers`
 */
function _getOrCreateSheet_(ss, name, headers, info) {
  const existing = ss.getSheetByName(name);
  if (existing) {
    const actual = existing.getRange(HEADER_ROW, 1, 1, headers.length).getValues()[0].map(String);
    if (actual.join("|") !== headers.join("|")) {
      throw new Error(
        `Tab "${name}" has unexpected headers: expected [${headers.join(", ")}], found [${actual.join(", ")}]`,
      );
    }
    return { sheet: existing, created: false };
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
  return { sheet: sheet, created: true };
}

/**
 * resize a tab to exactly the given number of rows and columns
 * adding or deleting rows and columns at the end as needed
 *
 * always keeps at least one row below any frozen rows, because Sheets does not
 * allow deleting every unfrozen row.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet    tab to resize
 * @param {number}                             rows     total rows wanted, including the info and header rows
 * @param {number}                             columns  total columns wanted
 */
function _fitSheet_(sheet, rows, columns) {
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
 * @param {number} column  1-based column number
 * @returns {string} column letters
 */
function _columnLetter_(column) {
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
 * @param {string} sheetName  tab name
 * @param {number} column     1-based column number
 * @returns {string} the reference, for use inside formulas
 */
function _columnBelowHeader_(sheetName, column) {
  const letter = _columnLetter_(column);
  return `'${sheetName}'!$${letter}$${DATA_ROW}:$${letter}`;
}

/**
 * compare two dates for sorting, oldest first
 * null (invalid folder name) sorts last
 *
 * @param {Date|null} a  first date
 * @param {Date|null} b  second date
 * @returns {number} negative if a comes first, positive if b comes first, 0 if equal
 */
function _compareDates_(a, b) {
  if (a === null || b === null) return (a === null ? 1 : 0) - (b === null ? 1 : 0);
  return a.getTime() - b.getTime();
}
