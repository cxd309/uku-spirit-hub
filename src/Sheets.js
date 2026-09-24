/**
 * get a tab by name, creating with a bold frozen header row if does not exist
 *
 * if the tab exists, header row must match `headers` exactly
 * code and sheet cannot silently diagree
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet}  ss      spreadsheet to examine
 * @param {string}                                    name    tab name
 * @param {readonly string[]}                         headers expected header row
 * @returns {{sheet: GoogleAppsScript.Spreadsheet.Sheet, created: boolean}} the tab and if created
 * @throws {Error} if the tab exists but its header row differs from `headers`
 */
function _getOrCreateSheet_(ss, name, headers) {
  const existing = ss.getSheetByName(name);
  if (existing) {
    const actual = existing.getRange(1, 1, 1, headers.length).getValues()[0].map(String);
    if (actual.join("|") !== headers.join("|")) {
      throw new Error(
        `Tab "${name}" has unexpectedd headers: expected [${headers.join(", ")}], found [${actual.join(", ")}]`,
      );
    }
    return { sheet: existing, created: false };
  }

  const sheet = ss.insertSheet(name);
  _fitSheet_(sheet, 2, headers.length);
  sheet.getRange(1, 1, 1, headers.length).setValues([[...headers]]).setFontWeight("bold");
  sheet.setFrozenRows(1);
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
 * @param {number}                             rows     total rows wanted, including the header
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
