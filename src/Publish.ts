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
 * a published file found in a SOTG Folder
 */
interface PublishedFile {
  /** Drive file ID */
  id: string;
  /** when it was last written */
  modifiedTime: Date;
}

/**
 * what looking in a SOTG Folder found: the files with the published name, or why the folder could not be read
 */
type FolderLookup = { ok: true; files: PublishedFile[] } | { ok: false; reason: string };

/**
 * a tournament's Publish Status, the link to its published file, and anything to say in Message
 * pure, no google calls
 *
 * NO FOLDER without a SOTG Folder, ERROR if the folder link cannot be read, NOT PUBLISHED without a file of
 * the published name, otherwise PUBLISHED, or OUT OF DATE when the Form Responses or schedule it was made from
 * changed after it was written
 *
 * @param event   the tournament
 * @param lookup  what looking in its SOTG Folder found, undefined when there is no folder
 * @returns the status, the file's link, "" if none, and the reason for an ERROR or a note, "" if none
 */
function _publishStatus_(
  event: EventRecord,
  lookup: FolderLookup | undefined,
): { status: string; url: string; reason: string } {
  if (event.sotgFolderUrl === "") return { status: PUBLISH_STATUS.NO_FOLDER, url: "", reason: "" };
  if (_folderIdFromUrl_(event.sotgFolderUrl) === "") {
    return { status: PUBLISH_STATUS.ERROR, url: "", reason: "SOTG Folder is not a Google Drive folder link" };
  }
  if (!lookup) return { status: PUBLISH_STATUS.ERROR, url: "", reason: "SOTG Folder not checked" };
  if (!lookup.ok) return { status: PUBLISH_STATUS.ERROR, url: "", reason: lookup.reason };
  if (lookup.files.length === 0) return { status: PUBLISH_STATUS.NOT_PUBLISHED, url: "", reason: "" };

  const file = lookup.files[0];
  const url = `https://docs.google.com/spreadsheets/d/${file.id}/edit`;
  const reason = lookup.files.length > 1
    ? `${lookup.files.length} files with the published name, the first is used`
    : "";
  const changed = [event.scheduleImportDate, event.spiritFormImportDate]
    .some((imported) => imported !== null && imported.getTime() > file.modifiedTime.getTime() + VERSION_TOLERANCE_MS);
  return { status: changed ? PUBLISH_STATUS.OUT_OF_DATE : PUBLISH_STATUS.PUBLISHED, url, reason };
}

/**
 * look for the published file in many SOTG Folders at once
 * only file names and dates are read, which the hub's Drive permission allows
 *
 * @param wanted  folder ID and the file name to look for in it
 * @returns one lookup per entry of `wanted`, in the same order
 */
function _lookUpPublishedFiles_(wanted: { folderId: string; name: string }[]): FolderLookup[] {
  if (wanted.length === 0) return [];
  const quote = (text: string) => text.replace(/\\/g, "\\\\").replace(/'/g, "\\'");
  const responses = _fetchGoogleJson_(
    wanted.map((w) =>
      "https://www.googleapis.com/drive/v3/files?q="
      + encodeURIComponent(`'${quote(w.folderId)}' in parents and name = '${quote(w.name)}' and trashed = false`)
      + `&fields=${encodeURIComponent("files(id,modifiedTime)")}`
      + "&orderBy=modifiedTime%20desc&supportsAllDrives=true&includeItemsFromAllDrives=true&corpora=allDrives"
    ),
  );
  return responses.map((response): FolderLookup => {
    if (response.code === 404) return { ok: false, reason: "SOTG Folder not found, or not shared with you" };
    if (response.code !== 200) return { ok: false, reason: `could not read SOTG Folder (${_apiError_(response)})` };
    const files = (response.body as { files?: { id?: string; modifiedTime?: string }[] }).files ?? [];
    return {
      ok: true,
      files: files.filter((f) => f.id).map((f) => ({
        id: f.id as string,
        modifiedTime: new Date(f.modifiedTime ?? 0),
      })),
    };
  });
}

/**
 * the events with their Publish Status, Results and Breakdown URL and Publish message brought up to date
 *
 * @param events   every event
 * @param pattern  the Published File Name setting
 * @returns the events, in the order given
 */
function _withPublishStatus_(events: EventRecord[], pattern: string): EventRecord[] {
  const checked = events.filter((e) => _folderIdFromUrl_(e.sotgFolderUrl) !== "");
  const lookups = _lookUpPublishedFiles_(
    checked.map((e) => ({ folderId: _folderIdFromUrl_(e.sotgFolderUrl), name: _publishedFileName_(pattern, e) })),
  );
  const lookupOf = new Map(checked.map((e, i) => [e, lookups[i]]));
  return events.map((event) => {
    const { status, url, reason } = _publishStatus_(event, lookupOf.get(event));
    return {
      ...event,
      publishStatus: status,
      publishedUrl: url,
      message: _sourceMessage_(event.message, "Publish", reason),
    };
  });
}
