/**
 * Outcome of importing one results file.
 */
type ImportResult = {
  ok: true;
  responses: ResponseRecord[];
  problems: RowProblem[];
  version: Date;
} | {
  ok: false;
  reason: string;
};

/**
 * the Sheets advanced service
 *
 * @returns the service
 * @throws {Error} if it is not turned on for this script
 */
function _sheetsService_(): GoogleAppsScript.Sheets {
  if (typeof Sheets === "undefined" || !Sheets) {
    throw new Error("The Google Sheets API service is not turned on: in the Apps Script editor add it under Services");
  }
  return Sheets;
}

/**
 * a cell as the Sheets service returns it, turned into the value SpreadsheetApp's getValues would give
 * numbers, text and ticks as themselves, errors as their text (e.g. "#N/A"), empty cells as ""
 * pure, no google calls
 *
 * @param cell  the cell, undefined when the row stops before it
 * @returns the value
 */
function _cellValue_(cell: GoogleAppsScript.Sheets.Schema.CellData | undefined): unknown {
  const value = cell?.effectiveValue;
  if (!value) return "";
  if (value.numberValue !== undefined) return value.numberValue;
  if (value.stringValue !== undefined) return value.stringValue;
  if (value.boolValue !== undefined) return value.boolValue;
  return cell?.formattedValue ?? "";
}

/**
 * every tab of a results file with all its values, in one call to the Sheets service
 * much faster than opening the file with SpreadsheetApp and reading tab by tab
 *
 * the service leaves out empty cells at the end of each row and empty rows at the end of the tab,
 * rows are padded so every row is as wide as the widest, like getDataRange().getValues()
 *
 * @param fileId  Drive file ID of the results file
 * @returns the tabs in file order
 */
function _readResultsTabs_(fileId: string): ResultsTab[] {
  const book = _sheetsService_().Spreadsheets.get(fileId, {
    includeGridData: true,
    fields: "sheets(properties(title),data(rowData(values(effectiveValue,formattedValue))))",
  });
  return (book.sheets ?? []).map((sheet) => {
    const rows = (sheet.data ?? []).flatMap((grid) => grid.rowData ?? []).map((row) =>
      (row.values ?? []).map(_cellValue_)
    );
    const width = Math.max(0, ...rows.map((row) => row.length));
    return {
      title: sheet.properties?.title ?? "",
      values: rows.map((row) => [...row, ...Array(width - row.length).fill("")]),
    };
  });
}

/**
 * read an event's results file into response records
 * uses only the file ID from the Tournaments tab, no folder scan
 *
 * any error from Google (file deleted, no permission) is caught and returned as a
 * reason, so one bad file cannot stop the whole refresh
 *
 * @param event  the event to import
 * @returns the responses read, or why the file could not be read
 */
function _importFile_(event: EventRecord): ImportResult {
  if (event.date === null) {
    return {
      ok: false,
      reason: "folder name must be \"YYYYMMDD Tournament name\", fix it then run Refresh Tournaments",
    };
  }
  if (!event.fileName.toLowerCase().includes(RESULTS_FILE_TEXT.toLowerCase())) {
    return {
      ok: false,
      reason:
        `file name "${event.fileName}" does not contain "${RESULTS_FILE_TEXT}", fix it then run Refresh Tournaments`,
    };
  }
  try {
    // version is read before the contents, so an edit made during the read is picked up next time
    const modified = _driveService_().Files.get(event.fileId, { fields: "modifiedTime", supportsAllDrives: true });
    const version = new Date(modified.modifiedTime ?? 0);
    const found = _findBreakdownTab_(_readResultsTabs_(event.fileId));
    if (!found.ok) return { ok: false, reason: found.reason };

    const { responses, problems } = _parseBreakdownRows_(found.tab.values, found.columns);
    return {
      ok: true,
      responses: responses.map((r) => ({
        fileId: event.fileId,
        sourceRow: r.sourceRow,
        scorer: r.scorer,
        receiver: r.receiver,
        ...r.scores,
        comment: r.comment,
      })),
      problems: problems,
      version: version,
    };
  } catch (e) {
    const message = e instanceof Error ? e.message : String(e);
    return { ok: false, reason: `could not open file (${message}), try Refresh Tournaments first` };
  }
}

/**
 * import every tournament marked NEW or REFRESH on the Tournaments tab
 * works from the Tournaments tab only, run Refresh Tournaments first to find new or edited files
 *
 * successful imports replace that event's responses and clear its status
 * failed imports mark the event ERROR and keep its previous responses
 * responses are written before Tournaments, so an interrupted run is simply repeated
 *
 * @returns a one-line summary of the refresh
 */
function _refreshResults_(): string {
  const eventsSheet = _getEventsSheet_();
  const responsesSheet = _getResponsesSheet_();
  const events = _readEvents_(eventsSheet);
  if (events.length === 0) return "No tournaments found: run Refresh Tournaments first";

  const toImport = events.filter((e) => e.status === EVENT_STATUS.NEW || e.status === EVENT_STATUS.REFRESH);
  if (toImport.length === 0) return "Nothing to refresh: no tournaments are NEW or REFRESH";

  const results: Map<string, ImportResult> = _timed_(
    "read files",
    () => new Map(toImport.map((e) => [e.fileId, _importFile_(e)])),
  );

  const updatedEvents = events.map((event) => {
    const result = results.get(event.fileId);
    if (!result) return event;
    if (!result.ok) return { ...event, status: EVENT_STATUS.ERROR, message: result.reason };
    const skipped = result.problems.length === 0
      ? ""
      : `${result.problems.length} row(s) skipped: ${
        result.problems.map((p) => `row ${p.sourceRow} (${p.reason})`).join("; ")
      }`;
    return { ...event, status: EVENT_STATUS.OK, message: skipped, importedVersion: result.version };
  });

  const replacedIds = new Set([...results].filter(([, r]) => r.ok).map(([id]) => id));
  const imported = [...results.values()].flatMap((r) => (r.ok ? r.responses : []));
  const responses = _sortResponses_([
    ..._readResponses_(responsesSheet).filter((r) => !replacedIds.has(r.fileId)),
    ...imported,
  ], updatedEvents);

  _timed_("write results", () => _writeResponses_(responsesSheet, responses));
  _timed_("write tournaments", () => _writeEvents_(eventsSheet, updatedEvents));

  const internationalIds = new Set(updatedEvents.filter((e) => e.international).map((e) => e.fileId));
  const teamsSheet = _getTeamsSheet_();
  const sortedEvents = _sortEvents_(updatedEvents);
  const teamClubs = _timed_("write teams", () =>
    _writeTeams_(
      teamsSheet,
      _mergeTeams_(_readTeams_(teamsSheet), _teamNamesFromResponses_(responses, internationalIds)),
      responses,
      sortedEvents,
    ));
  _timed_("rebuild clubs", () => _rebuildClubs_(teamClubs, responses, sortedEvents));

  const failed = [...results.values()].filter((r) => !r.ok).length;
  return `Imported ${results.size - failed} tournament(s), ${imported.length} response(s)`
    + (failed > 0 ? `; ${failed} failed (see Tournaments tab)` : "");
}
