/**
 * placeholder in the Published File Name setting, replaced by the tournament's Event Name
 */
const PUBLISHED_NAME_PLACEHOLDER = "{event_name}";

/**
 * the name of a tournament's published Results and Breakdown file
 * pure, no google calls
 *
 * @param pattern  the Published File Name setting
 * @param event    the tournament
 * @returns the name, e.g. "NUXIR 2026-27 Spirit Results and Breakdown"
 */
function _publishedFileName_(pattern: string, event: EventRecord): string {
  return pattern.split(PUBLISHED_NAME_PLACEHOLDER).join(event.eventName).trim();
}

/**
 * the folder ID inside a Google Drive folder link
 * pure, no google calls
 *
 * @param url  e.g. https://drive.google.com/drive/folders/<id>?usp=sharing
 * @returns the ID, "" if the link is not a Drive folder link
 */
function _folderIdFromUrl_(url: string): string {
  const match = /\/folders\/([a-zA-Z0-9_-]+)/.exec(url);
  return match ? match[1] : "";
}

/**
 * Drive app property the hub puts on files it publishes, holding its own spreadsheet ID
 * only the app that set it can read it, so a file without it was made somewhere else
 */
const PUBLISHED_BY_PROPERTY = "spiritHub";

/**
 * a published Results and Breakdown file
 */
interface PublishedFile {
  /** Drive file ID */
  id: string;
  /** file name */
  name: string;
  /** when it was last written */
  modifiedTime: Date;
  /** whether this hub created it */
  own: boolean;
}

/**
 * what looking in a SOTG Folder found: the files with the published name, or why the folder could not be read
 */
type FolderLookup = { ok: true; files: PublishedFile[] } | { ok: false; reason: string };

/**
 * a tournament's Publish Status, the link to its published file, and anything to say in Message
 * pure, no google calls
 *
 * this hub's own file comes first: one it created, with the published name, linked or in the SOTG Folder.
 * It is PUBLISHED, or OUT OF DATE when the Form Responses or schedule it was made from changed after it was written
 * otherwise a file the Schedule Manager links to, whatever its name, or one in the SOTG Folder with the published name,
 * is PUBLISHED EXTERNALLY
 * without a file: NO FOLDER without a SOTG Folder, ERROR if it cannot be read, otherwise NOT PUBLISHED
 *
 * @param event   the tournament
 * @param name    the published file name
 * @param linked  files the Schedule Manager or the last check linked to, that could be read
 * @param lookup  what looking in its SOTG Folder found, undefined when it was not looked in
 * @returns the status, the file's link, "" if none, and the reason for an ERROR or a note, "" if none
 */
function _publishStatus_(
  event: EventRecord,
  name: string,
  linked: PublishedFile[],
  lookup: FolderLookup | undefined,
): { status: string; url: string; reason: string } {
  const inFolder = lookup && lookup.ok ? lookup.files : [];
  const urlOf = (file: PublishedFile) => `https://docs.google.com/spreadsheets/d/${file.id}/edit`;

  const own = [...linked, ...inFolder].find((f) => f.own && f.name.trim() === name);
  if (own) {
    const others = inFolder.filter((f) => f.id !== own.id).length;
    const reason = others > 0 ? `${others + 1} files with the published name, this hub's is used` : "";
    const modified = own.modifiedTime.getTime();
    const changed = [event.scheduleImportDate, event.spiritFormImportDate]
      .some((imported) => imported !== null && imported.getTime() > modified + VERSION_TOLERANCE_MS);
    return { status: changed ? PUBLISH_STATUS.OUT_OF_DATE : PUBLISH_STATUS.PUBLISHED, url: urlOf(own), reason };
  }

  const external = linked.find((f) => !f.own) ?? inFolder.find((f) => !f.own);
  if (external) return { status: PUBLISH_STATUS.PUBLISHED_EXTERNALLY, url: urlOf(external), reason: "" };

  if (event.sotgFolderUrl === "") return { status: PUBLISH_STATUS.NO_FOLDER, url: "", reason: "" };
  if (_folderIdFromUrl_(event.sotgFolderUrl) === "") {
    return { status: PUBLISH_STATUS.ERROR, url: "", reason: "SOTG Folder is not a Google Drive folder link" };
  }
  if (!lookup) return { status: PUBLISH_STATUS.ERROR, url: "", reason: "SOTG Folder not checked" };
  if (!lookup.ok) return { status: PUBLISH_STATUS.ERROR, url: "", reason: lookup.reason };
  return { status: PUBLISH_STATUS.NOT_PUBLISHED, url: "", reason: "" };
}

/**
 * a Drive file's details as the API returns them
 */
type DriveFileDetails = {
  id?: string;
  name?: string;
  modifiedTime?: string;
  trashed?: boolean;
  appProperties?: Record<string, string>;
};

/**
 * the fields asked of Drive for a published file
 */
const PUBLISHED_FILE_FIELDS = "id,name,modifiedTime,trashed,appProperties";

/**
 * a file from Drive's details, and whether this hub created it
 *
 * @param details  as the Drive API returns them
 * @param hubId    this hub's spreadsheet ID
 * @returns the file
 */
function _publishedFile_(details: DriveFileDetails, hubId: string): PublishedFile {
  return {
    id: details.id ?? "",
    name: details.name ?? "",
    modifiedTime: new Date(details.modifiedTime ?? 0),
    own: details.appProperties?.[PUBLISHED_BY_PROPERTY] === hubId,
  };
}

/**
 * read linked files, as the Schedule Manager or the last check gave them
 * only file names and dates are read, which the hub's Drive permission allows
 *
 * @param urls   spreadsheet links
 * @param hubId  this hub's spreadsheet ID
 * @returns one file per link, in the same order, null if it is not a link, cannot be read or is in the bin
 */
function _readLinkedFiles_(urls: string[], hubId: string): (PublishedFile | null)[] {
  const ids = urls.map((url) => _spreadsheetIdFromUrl_(url));
  const toFetch = ids.flatMap((id, i) => (id === "" ? [] : [i]));
  const responses = _fetchGoogleJson_(
    toFetch.map((i) =>
      `https://www.googleapis.com/drive/v3/files/${encodeURIComponent(ids[i])}`
      + `?fields=${encodeURIComponent(PUBLISHED_FILE_FIELDS)}&supportsAllDrives=true`
    ),
  );
  const files: (PublishedFile | null)[] = urls.map(() => null);
  toFetch.forEach((i, j) => {
    const details = responses[j].body as DriveFileDetails;
    if (responses[j].code === 200 && !details.trashed) files[i] = _publishedFile_(details, hubId);
  });
  return files;
}

/**
 * look for the published file in many SOTG Folders at once, newest first
 * only file names and dates are read, which the hub's Drive permission allows
 *
 * @param wanted  folder ID and the file name to look for in it
 * @param hubId   this hub's spreadsheet ID
 * @returns one lookup per entry of `wanted`, in the same order
 */
function _lookUpPublishedFiles_(wanted: { folderId: string; name: string }[], hubId: string): FolderLookup[] {
  if (wanted.length === 0) return [];
  const quote = (text: string) => text.replace(/\\/g, "\\\\").replace(/'/g, "\\'");
  const responses = _fetchGoogleJson_(
    wanted.map((w) =>
      "https://www.googleapis.com/drive/v3/files?q="
      + encodeURIComponent(`'${quote(w.folderId)}' in parents and name = '${quote(w.name)}' and trashed = false`)
      + `&fields=${encodeURIComponent(`files(${PUBLISHED_FILE_FIELDS})`)}`
      + "&orderBy=modifiedTime%20desc&supportsAllDrives=true&includeItemsFromAllDrives=true&corpora=allDrives"
    ),
  );
  return responses.map((response): FolderLookup => {
    if (response.code === 404) return { ok: false, reason: "SOTG Folder not found, or not shared with you" };
    if (response.code !== 200) return { ok: false, reason: `could not read SOTG Folder (${_apiError_(response)})` };
    const files = (response.body as { files?: DriveFileDetails[] }).files ?? [];
    return { ok: true, files: files.filter((f) => f.id).map((f) => _publishedFile_(f, hubId)) };
  });
}

/**
 * the events with their Publish Status, Results and Breakdown URL and Publish message brought up to date
 * reads the files linked from the Schedule Manager and from the last check, and looks in every SOTG Folder
 *
 * @param events   every event
 * @param pattern  the Published File Name setting
 * @param links    Spirit Form URL → the Schedule List's SOTG Results and Breakdown link, empty when not read
 * @returns the events, in the order given
 */
function _withPublishStatus_(events: EventRecord[], pattern: string, links: Map<string, string>): EventRecord[] {
  const hubId = SpreadsheetApp.getActiveSpreadsheet().getId();
  const nameOf = (e: EventRecord) => _publishedFileName_(pattern, e);
  const linksOf = (e: EventRecord) => [...new Set([links.get(e.spiritFormUrl) ?? "", e.publishedUrl])].filter((u) => u);

  const urls = [...new Set(events.flatMap(linksOf))];
  const read = _readLinkedFiles_(urls, hubId);
  const fileAt = new Map(urls.map((url, i) => [url, read[i]]));

  const searched = events.filter((e) => _folderIdFromUrl_(e.sotgFolderUrl) !== "");
  const lookups = _lookUpPublishedFiles_(
    searched.map((e) => ({ folderId: _folderIdFromUrl_(e.sotgFolderUrl), name: nameOf(e) })),
    hubId,
  );
  const lookupOf = new Map(searched.map((e, i) => [e, lookups[i]]));

  return events.map((event) => {
    const linked = linksOf(event).flatMap((url) => {
      const file = fileAt.get(url);
      return file ? [file] : [];
    });
    const { status, url, reason } = _publishStatus_(event, nameOf(event), linked, lookupOf.get(event));
    return {
      ...event,
      publishStatus: status,
      publishedUrl: url,
      message: _sourceMessage_(event.message, "Publish", reason),
    };
  });
}
