/**
 * text every results file name must contain
 * matched case-insensitively
 * Google Sheets without it are still listed on the Tournaments tab, as ERROR, so they are not missed
 */
const RESULTS_FILE_TEXT = "Spirit Results and Breakdown";

/**
 * A Google Sheet found under the category folder, expected to be a results file
 */
interface ResultsFile {
  /** drive file ID */
  id: string;
  /** file name, e.g. "WWUXIR Spirit Results and Breakdown" */
  name: string;
  /** drive ID of the folder containing file */
  folderId: string;
  /** name of folder containing the file (default tournament name) */
  folderName: string;
  /** top-level folder, e.g. University, Club, Youth */
  category: string;
  /** folder names between the category and the file */
  path: string[];
  /** when the file was last modified in drive */
  lastUpdated: Date;
}

/**
 * find the season folder: the Drive folder containing this hub spreadsheet
 *
 * the hub must sit directly inside the season folder (e.g. '2025-26/')
 *
 * @returns the season folder
 * @throws {Error} if the hub spreadsheet is not inside a folder
 */
function _findSeasonFolder_(): GoogleAppsScript.Drive.Folder {
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
 * @param category name of the folder directly inside season folder
 * @returns all results files found
 * @throws {Error} if the category folder doesn't exist
 */
function _scanCategory_(category: string): ResultsFile[] {
  const season = _findSeasonFolder_();
  const matches = season.getFoldersByName(category);
  if (!matches.hasNext()) {
    const available: string[] = [];
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
 * walk a folder recursively and collect every Google Sheet
 *
 * sheets are collected whatever their name, the Tournaments tab marks the badly named ones ERROR
 *
 * @param folder    folder to search
 * @param category  category the folder belongs to
 * @param path      each folder name on the path from category to folder
 * @returns results files in this folder and everything below it
 */
function _scanFolder_(folder: GoogleAppsScript.Drive.Folder, category: string, path: string[]): ResultsFile[] {
  let found = [];

  const files = folder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    if (file.getMimeType() !== MimeType.GOOGLE_SHEETS) continue;

    found.push({
      id: file.getId(),
      name: file.getName(),
      folderId: folder.getId(),
      folderName: folder.getName(),
      category: category,
      path: path,
      lastUpdated: new Date(file.getLastUpdated().getTime()),
    });
  }

  const subfolders = folder.getFolders();
  while (subfolders.hasNext()) {
    const sub = subfolders.next();
    found = found.concat(_scanFolder_(sub, category, [...path, sub.getName()]));
  }
  return found;
}
