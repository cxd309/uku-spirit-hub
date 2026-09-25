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
 *
 * any error from Google (file deleted, no permission) is caught and returned as a
 * reason, so one bad file cannot stop the whole import
 *
 * @param {EventRecord} event  the event to import
 * @param {ResultsFile} file   the event's file, from the current scan
 * @returns {ImportResult} the responses read, or why the file could not be read
 */
function _importFile_(event, file) {
  try {
    const folder = _parseFolderName_(file.folderName);
    if (!folder.ok) return { ok: false, reason: folder.reason };

    const found = _findBreakdownSheet_(SpreadsheetApp.openById(file.id));
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
      version: file.lastUpdated,
    };
  } catch (e) {
    return { ok: false, reason: `could not read file: ${e instanceof Error ? e.message : String(e)}` };
  }
}

/**
 * scan the category folder, then import every event marked NEW or REFRESH
 *
 * successful imports replace that event's responses and clear its status
 * failed imports mark the event ERROR and keep its previous responses
 * responses are written before Events, so an interrupted run is simply repeated
 * @returns {string} a one-line summary of the import
 */
function _importEvents_() {
  const eventsSheet = _getEventsSheet_();
  const responsesSheet = _getResponsesSheet_();

  const files = _scanCategory_(_readConfig_().category);
  const fileById = new Map(files.map((f) => [f.id, f]));
  const events = _syncEvents_(_readEvents_(eventsSheet), files);

  /** @type {Map<string, ImportResult>} */
  const results = new Map();
  for (const event of events) {
    const file = fileById.get(event.fileId);
    if (file && (event.status === EVENT_STATUS.NEW || event.status === EVENT_STATUS.REFRESH)) {
      results.set(event.fileId, _importFile_(event, file));
    }
  }

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
  const responses = [..._readResponses_(responsesSheet).filter((r) => !replacedIds.has(r.fileId)), ...imported];

  _writeResponses_(responsesSheet, responses);
  _writeEvents_(eventsSheet, updatedEvents);

  _getNameRulesSheet_();
  const internationalIds = new Set(updatedEvents.filter((e) => e.international).map((e) => e.fileId));
  const teamsSheet = _getTeamsSheet_();
  _writeTeams_(
    teamsSheet,
    _mergeTeams_(_readTeams_(teamsSheet), _teamNamesFromResponses_(responses, internationalIds)),
  );
  SpreadsheetApp.flush();
  _rebuildClubs_();

  const failed = [...results.values()].filter((r) => !r.ok).length;
  const imported_ = results.size - failed;
  if (results.size === 0) return "Nothing to import: no events are NEW or REFRESH";
  return `Imported ${imported_} event(s), ${imported.length} response(s)`
    + (failed > 0 ? `; ${failed} failed (see Events tab)` : "");
}
