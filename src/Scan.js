/**
 * file-name suffix that marks a GoogleSheet as a results file
 * matched case-insensitively, ignore trailing whitespace
 */
const RESULTS_FILE_SUFFIX = "spirit results and breakdown";

/**
 * A results file found under the season folder
 *
 * @typedef {Object} ResultsFile
 * @property {string}   id          drive file ID
 * @property {string}   name        file name, e.g. "WWUXIR Spirit Results and Breakdown"
 * @property {string}   folderId    drive ID of the folder containing file
 * @property {string}   folderName  name of folder containing the file (default tournament name)
 * @property {string}   category    top-level folder, e.g. University, Club, Youth
 * @property {string[]} path        folder names between the category and the file
 * @property {Date}     lastUpdated when the file was last modified in drive
 */

/**
 * find the season folder: the Drive folder containing this hub spreadsheet
 *
 * the hub must sit directly inside the season folder (e.g. '2025-26/')
 *
 * @returns {GoogleAppsScript.Drive.Folder} the season folder
 * @throws {Error} if the hub spreadsheet is not inside a folder
 */
function _findSeasonFolder_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const file = DriveApp.getFileById(ss.getId());
  const parents = file.getParents();
  if (!parents.hasNext()) {
    throw new Error("Hub spreadsheet is not inside a folder");
  }
  return parents.next();
}

/**
 * find every results file in the season folder for this category
 *
 * a results file is any Google Sheet inside a category folder, at any depth
 * files directly in the season folder (like the Hub itself) are ignored
 * @param {string} category name of the folder directly inside season folder
 * @returns {ResultsFile[]}   all results files found
 * @throws {Error} if the category folder doesn't exist
 */
function _scanCategory_(category) {
  const season = _findSeasonFolder_();
  const matches = season.getFoldersByName(category);
  if (!matches.hasNext()) {
    /** @type {string[]} */
    const available = [];
    const folders = season.getFolders();
    while (folders.hasNext()) available.push(folders.next().getName());
    throw new Error(`No "${category}" folder in "${season.getName()}". Available: ${available.join(", ")}`);
  }
  const folder = matches.next();
  if (matches.hasNext()) {
    throw new Error(`More than one "${category}" folder in "${season.getName()}"`);
  }
  return _scanFolder_(folder, category, []);
}

/**
 * walk a folder recursively and collect all results files
 *
 * google sheets without results suffix are logged as warnings
 *
 * @param {GoogleAppsScript.Drive.Folder} folder    folder to search
 * @param {string}                        category  category the folder belongs to
 * @param {string[]}                      path      each folder name on the path from category to folder
 * @returns {ResultsFile[]}                         results files in this folder and everything below it
 */
function _scanFolder_(folder, category, path) {
  let found = [];

  const files = folder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    if (file.getMimeType() !== MimeType.GOOGLE_SHEETS) continue;

    const name = file.getName();
    if (name.trim().toLowerCase().endsWith(RESULTS_FILE_SUFFIX)) {
      found.push({
        id: file.getId(),
        name: name,
        folderId: folder.getId(),
        folderName: folder.getName(),
        category: category,
        path: path,
        lastUpdated: new Date(file.getLastUpdated().getTime()),
      });
    } else {
      console.warn(`Skipping Google Sheet without the results suffix: ${[category, ...path, name].join(" / ")}`);
    }
  }

  const subfolders = folder.getFolders();
  while (subfolders.hasNext()) {
    const sub = subfolders.next();
    found = found.concat(_scanFolder_(sub, category, [...path, sub.getName()]));
  }
  return found;
}
