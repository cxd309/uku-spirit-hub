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
 * Drive's type for a Google Sheet, given when a published file is made
 */
const DRIVE_SHEET_MIME_TYPE = "application/vnd.google-apps.spreadsheet";

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

/**
 * Publish Statuses that Publish Results writes a file for
 * PUBLISHED EXTERNALLY gets a file of this hub's own, unless one by another hand already has the published name
 */
const PUBLISHABLE_STATUSES: readonly string[] = Object.freeze([
  PUBLISH_STATUS.NOT_PUBLISHED,
  PUBLISH_STATUS.OUT_OF_DATE,
  PUBLISH_STATUS.PUBLISHED_EXTERNALLY,
]);

/**
 * how one tournament's publish went
 */
type PublishOutcome = { ok: true; fileId: string; created: boolean } | { ok: false; reason: string };

/**
 * the published file a tournament should be written to, from what is in its SOTG Folder
 * pure, no google calls
 *
 * this hub's own file with the published name is updated, a file with that name made by another hand blocks it,
 * as this hub cannot change it, otherwise a new file is made
 *
 * @param event   the tournament
 * @param name    the published file name
 * @param lookup  the files in its SOTG Folder with that name
 * @returns the file ID to update, "" to make a new one, or why it cannot be published
 */
function _publishTarget_(
  event: EventRecord,
  name: string,
  lookup: FolderLookup,
): { ok: true; fileId: string } | { ok: false; reason: string } {
  if (!lookup.ok) return { ok: false, reason: lookup.reason };
  const own = lookup.files.find((f) => f.own);
  if (own) return { ok: true, fileId: own.id };
  if (lookup.files.length > 0) {
    return { ok: false, reason: `"${name}" in the SOTG Folder was not made by this hub, rename or remove it first` };
  }
  if (event.publishStatus === PUBLISH_STATUS.OUT_OF_DATE) {
    const linked = _spreadsheetIdFromUrl_(event.publishedUrl);
    if (linked !== "") return { ok: true, fileId: linked };
  }
  return { ok: true, fileId: "" };
}

/**
 * the tabs of a published file the requests are written for: the first two, or a new second one
 * pure, no google calls
 *
 * @param sheets  the file's tabs, as the Sheets API gives them
 * @returns the Results and Breakdown tabs
 */
function _publishTabs_(
  sheets: GoogleAppsScript.Sheets.Schema.Sheet[],
): { results: PublishedTab; breakdown: PublishedTab } {
  const ordered = [...sheets].sort((a, b) => (a.properties?.index ?? 0) - (b.properties?.index ?? 0));
  const tab = (sheet: GoogleAppsScript.Sheets.Schema.Sheet | undefined, fallbackId: number): PublishedTab =>
    sheet
      ? { sheetId: sheet.properties?.sheetId ?? 0, isNew: false, ruleCount: sheet.conditionalFormats?.length ?? 0 }
      : { sheetId: fallbackId, isNew: true, ruleCount: 0 };
  const nextId = Math.max(0, ...sheets.map((s) => s.properties?.sheetId ?? 0)) + 1;
  return { results: tab(ordered[0], nextId), breakdown: tab(ordered[1], nextId + 1) };
}

/**
 * write the Results and Breakdown file of every tournament ready to publish, then check Publish Status again
 * ready: Include Results ticked, Event Status OK, a SOTG Folder, and NOT PUBLISHED, OUT OF DATE or PUBLISHED EXTERNALLY
 * a new file is made in the SOTG Folder with the published name and this hub's label, an existing one of this hub's
 * is rewritten in place so its link never changes, files made by another hand are never touched
 * the content is today's layout, from the counted forms on Results
 *
 * @returns a one-line summary
 */
function _publishResults_(): string {
  const settings = _readHubSettings_();
  const sheet = _getEventsSheet_();
  const events = _readEvents_(sheet);
  const responses = _readResponses_(_getResponsesSheet_());
  const hubId = SpreadsheetApp.getActiveSpreadsheet().getId();
  const nameOf = (e: EventRecord) => _publishedFileName_(settings.publishedFileName, e);

  const ready = events.filter((e) =>
    e.includeResults
    && e.eventStatus === SOURCE_STATUS.OK
    && _folderIdFromUrl_(e.sotgFolderUrl) !== ""
    && PUBLISHABLE_STATUSES.includes(e.publishStatus)
  );
  if (ready.length === 0) {
    return "Nothing to publish: no included tournament with Event Status OK, an SOTG Folder, "
      + "and Publish Status NOT PUBLISHED, OUT OF DATE or PUBLISHED EXTERNALLY";
  }

  const outcomes: Map<EventRecord, PublishOutcome> = new Map();
  const lookups = _timed_(
    "check publish folders",
    () =>
      _lookUpPublishedFiles_(
        ready.map((e) => ({ folderId: _folderIdFromUrl_(e.sotgFolderUrl), name: nameOf(e) })),
        hubId,
      ),
  );
  const targets: Map<EventRecord, string> = new Map();
  ready.forEach((event, i) => {
    const target = _publishTarget_(event, nameOf(event), lookups[i]);
    if (target.ok) targets.set(event, target.fileId);
    else outcomes.set(event, target);
  });

  // make the files that do not exist yet, in each SOTG Folder, with this hub's label
  const toCreate = [...targets].filter(([, fileId]) => fileId === "").map(([event]) => event);
  const created = _timed_("create published files", () =>
    _sendGoogleJson_(toCreate.map((event) => ({
      url: "https://www.googleapis.com/drive/v3/files?supportsAllDrives=true&fields=id",
      body: {
        name: nameOf(event),
        mimeType: DRIVE_SHEET_MIME_TYPE,
        parents: [_folderIdFromUrl_(event.sotgFolderUrl)],
        appProperties: { [PUBLISHED_BY_PROPERTY]: hubId },
      },
    }))));
  const isNew: Set<EventRecord> = new Set();
  toCreate.forEach((event, i) => {
    const id = (created[i].body as { id?: string }).id;
    if (created[i].code === 200 && id) {
      targets.set(event, id);
      isNew.add(event);
    } else {
      targets.delete(event);
      outcomes.set(event, { ok: false, reason: `could not create the file (${_apiError_(created[i])})` });
    }
  });

  // read each file's tabs, then write both in one batch per file
  const writing = [...targets.keys()];
  const books = _timed_(
    "read published files",
    () =>
      _fetchGoogleJson_(
        writing.map((event) =>
          `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(targets.get(event) as string)}`
          + `?fields=${encodeURIComponent("sheets(properties(sheetId,index),conditionalFormats(booleanRule))")}`
        ),
      ),
  );
  const batches: { event: EventRecord; url: string; body: unknown }[] = [];
  writing.forEach((event, i) => {
    const fileId = targets.get(event) as string;
    if (books[i].code !== 200) {
      outcomes.set(event, { ok: false, reason: `could not open the published file (${_apiError_(books[i])})` });
      return;
    }
    const tabs = _publishTabs_((books[i].body as GoogleAppsScript.Sheets.Schema.Spreadsheet).sheets ?? []);
    const tables = _publishedTables_(responses.filter((r) => r.spiritFormUrl === event.spiritFormUrl));
    const titles = { results: `${event.eventName} SOTG Results`, breakdown: `${event.eventName} SOTG Breakdown` };
    batches.push({
      event,
      url: `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(fileId)}:batchUpdate`,
      body: { requests: _publishRequests_(tables, titles, tabs.results, tabs.breakdown) },
    });
  });
  const written = _timed_(
    "write published files",
    () => _sendGoogleJson_(batches.map(({ url, body }) => ({ url, body }))),
  );
  batches.forEach(({ event }, i) => {
    outcomes.set(
      event,
      written[i].code === 200
        ? { ok: true, fileId: targets.get(event) as string, created: isNew.has(event) }
        : { ok: false, reason: `could not write the published file (${_apiError_(written[i])})` },
    );
  });

  const updated = events.map((event) => {
    const outcome = outcomes.get(event);
    if (!outcome) return event;
    return outcome.ok
      ? {
        ...event,
        publishedUrl: `https://docs.google.com/spreadsheets/d/${outcome.fileId}/edit`,
        message: _sourceMessage_(event.message, "Publish", ""),
      }
      : { ...event, message: _sourceMessage_(event.message, "Publish", outcome.reason) };
  });
  const checked = _timed_(
    "check published files",
    () => _withPublishStatus_(updated, settings.publishedFileName, new Map()),
  );
  // a failed publish keeps its reason, the check only knows about the files it finds
  const final = checked.map((event, i) => {
    const outcome = outcomes.get(events[i]);
    return outcome && !outcome.ok
      ? { ...event, message: _sourceMessage_(event.message, "Publish", outcome.reason) }
      : event;
  });
  _timed_("write tournaments", () => _writeEvents_(sheet, final));

  const results = [...outcomes.values()];
  const madeCount = results.filter((o) => o.ok && o.created).length;
  const updatedCount = results.filter((o) => o.ok && !o.created).length;
  const failed = results.filter((o) => !o.ok).length;
  return `Published ${madeCount} new file(s), updated ${updatedCount}`
    + (failed > 0 ? `; ${failed} could not be published, see Message on Tournaments` : "");
}
