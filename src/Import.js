/**
 * Outcome of importing one results file.
 *
 * @typedef {{
 *  ok: true,
 *  responses: ResponseRecord[],
 *  problems: RowProblem[],
 *  version: Date
 * } | {
 *  ok: false,
 *  reason: string
 * }} ImportResult
 */

/**
 * read an event's results file into response records
 * uses only the file ID from the Tournaments tab, no folder scan
 *
 * any error from Google (file deleted, no permission) is caught and returned as a
 * reason, so one bad file cannot stop the whole refresh
 *
 * @param {EventRecord} event  the event to import
 * @returns {ImportResult} the responses read, or why the file could not be read
 */
function _importFile_(event) {
  if (event.date === null) {
    return {
      ok: false,
      reason: "folder name must be \"YYYYMMDD Tournament name\", fix it then run Refresh Tournaments",
    };
  }
  try {
    // version is read before the contents, so an edit made during the read is picked up next time
    const version = new Date(DriveApp.getFileById(event.fileId).getLastUpdated().getTime());
    const found = _findBreakdownSheet_(SpreadsheetApp.openById(event.fileId));
    if (!found.ok) return { ok: false, reason: found.reason };

    const { responses, problems } = _parseBreakdownRows_(found.sheet.getDataRange().getValues(), found.columns);
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
 * import every tournament marked NEW or REFRESH on the Tournaments tab, then refresh issues
 * works from the Tournaments tab only, run Refresh Tournaments first to find new or edited files
 *
 * successful imports replace that event's responses and clear its status
 * failed imports mark the event ERROR and keep its previous responses
 * responses are written before Tournaments, so an interrupted run is simply repeated
 *
 * @returns {string} a one-line summary of the refresh
 */
function _refreshResults_() {
  const eventsSheet = _getEventsSheet_();
  const responsesSheet = _getResponsesSheet_();
  const events = _readEvents_(eventsSheet);
  if (events.length === 0) return "No tournaments found: run Refresh Tournaments first";

  const toImport = events.filter((e) => e.status === EVENT_STATUS.NEW || e.status === EVENT_STATUS.REFRESH);
  if (toImport.length === 0) return "Nothing to refresh: no tournaments are NEW or REFRESH";

  /** @type {Map<string, ImportResult>} */
  const results = _timed_("read files", () => new Map(toImport.map((e) => [e.fileId, _importFile_(e)])));

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

  _getNameRulesSheet_();
  const internationalIds = new Set(updatedEvents.filter((e) => e.international).map((e) => e.fileId));
  const teamsSheet = _getTeamsSheet_();
  _timed_("write teams", () =>
    _writeTeams_(
      teamsSheet,
      _mergeTeams_(_readTeams_(teamsSheet), _teamNamesFromResponses_(responses, internationalIds)),
    ));
  _timed_("rebuild clubs", () => {
    SpreadsheetApp.flush();
    _rebuildClubs_();
  });
  const newIssues = _timed_("issues", () => _appendNewIssues_(responses, updatedEvents));

  const failed = [...results.values()].filter((r) => !r.ok).length;
  return `Imported ${results.size - failed} tournament(s), ${imported.length} response(s)`
    + (failed > 0 ? `; ${failed} failed (see Tournaments tab)` : "")
    + `; ${newIssues} new issue(s)`;
}
