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
  return _scanFolder_(folder, category);
}

/**
 * Drive types the scan looks for
 */
const DRIVE_FOLDER_TYPE = "application/vnd.google-apps.folder";
const DRIVE_SHEET_TYPE = "application/vnd.google-apps.spreadsheet";

/**
 * how many folders are asked about in one Drive query
 * keeps each query well under Drive's length limit
 */
const DRIVE_PARENTS_PER_QUERY = 40;

/**
 * the Drive advanced service
 *
 * @returns the service
 * @throws {Error} if it is not turned on for this script
 */
function _driveService_(): GoogleAppsScript.Drive {
  if (typeof Drive === "undefined" || !Drive) {
    throw new Error("The Drive API service is not turned on: in the Apps Script editor add it under Services");
  }
  return Drive;
}

/**
 * collect every Google Sheet in a folder and all the folders below it
 *
 * sheets are collected whatever their name, the Tournaments tab marks the badly named ones ERROR
 * works one level of folders at a time, asking Drive about many folders in one query
 * so a whole category takes a few calls rather than several per file
 * shared drives are included
 *
 * @param folder    the category folder
 * @param category  category the folder belongs to
 * @returns results files in this folder and everything below it
 */
function _scanFolder_(folder: GoogleAppsScript.Drive.Folder, category: string): ResultsFile[] {
  const drive = _driveService_();
  const found: ResultsFile[] = [];
  const seen: Set<string> = new Set();
  // folders to look in next: id → its name and the folder names from the category down to it
  let level: Map<string, { name: string; path: string[] }> = new Map([[folder.getId(), {
    name: folder.getName(),
    path: [],
  }]]);

  while (level.size > 0) {
    const next: Map<string, { name: string; path: string[] }> = new Map();
    const ids = [...level.keys()];
    for (let i = 0; i < ids.length; i += DRIVE_PARENTS_PER_QUERY) {
      const parents = ids.slice(i, i + DRIVE_PARENTS_PER_QUERY).map((id) => `'${id}' in parents`).join(" or ");
      const query = `(${parents}) and trashed = false `
        + `and (mimeType = '${DRIVE_SHEET_TYPE}' or mimeType = '${DRIVE_FOLDER_TYPE}')`;
      let pageToken = "";
      do {
        const page = drive.Files.list({
          q: query,
          fields: "nextPageToken, files(id, name, mimeType, modifiedTime, parents)",
          pageSize: 1000,
          corpora: "allDrives",
          includeItemsFromAllDrives: true,
          supportsAllDrives: true,
          ...(pageToken ? { pageToken } : {}),
        });
        for (const file of page.files ?? []) {
          const parentId = (file.parents ?? []).find((id) => level.has(id));
          const parent = parentId === undefined ? undefined : level.get(parentId);
          if (parentId === undefined || !parent || !file.id || !file.name || seen.has(file.id)) continue;
          seen.add(file.id);
          if (file.mimeType === DRIVE_FOLDER_TYPE) {
            next.set(file.id, { name: file.name, path: [...parent.path, file.name] });
          } else {
            found.push({
              id: file.id,
              name: file.name,
              folderId: parentId,
              folderName: parent.name,
              category: category,
              path: parent.path,
              lastUpdated: new Date(file.modifiedTime ?? 0),
            });
          }
        }
        pageToken = page.nextPageToken ?? "";
      } while (pageToken);
    }
    level = next;
  }
  return found;
}
