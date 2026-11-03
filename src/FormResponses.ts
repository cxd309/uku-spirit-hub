/**
 * exact header text of each column the hub reads from the SOTG form's responses
 * keys are names used in code
 * values must match row 1 exactly (after trimming), other columns are ignored
 */
const FORM_HEADERS = Object.freeze({
  timestamp: "Timestamp",
  yourTeam: "Your Team Name",
  opponentTeam: "Opponent Team Name",
  rules: "Opposition Rules Knowledge and Use",
  fouls: "Opposition Fouls and Body Contact",
  fairMindedness: "Opposition Fair-Mindedness",
  attitude: "Opposition Positive Attitude and Self-Control",
  communication: "Opposition Communication",
  comments: "Comments",
});

/**
 * header of the extra column on the Deleted forms tab, written by the Schedule Manager's delete step
 */
const DELETED_REASON_HEADER = "Reason for deletion";

/**
 * keys of the five spirit score columns in form order
 */
const SCORE_KEYS = Object.freeze(
  ["rules", "fouls", "fairMindedness", "attitude", "communication"] as const,
);

/**
 * 0-based column index of each form column, keyed like FORM_HEADERS
 */
type FormColumns = Record<keyof typeof FORM_HEADERS, number>;

/**
 * one tab of a Form Responses file, as read from the file
 */
interface ResultsTab {
  /** tab name */
  title: string;
  /** every value on the tab from row 1, each row padded to the same width */
  values: unknown[][];
}

/**
 * a tab with the form's headers and where its columns are
 */
interface FormTab {
  /** the tab */
  tab: ResultsTab;
  /** where each form column is */
  columns: FormColumns;
}

/**
 * the two tabs of a Form Responses file the hub reads, or why they could not be found
 */
type FormTabsLookup =
  | { ok: true; responses: FormTab; deleted: (FormTab & { reasonColumn: number }) | null }
  | { ok: false; reason: string };

/**
 * where each form column is in a header row
 * pure, no google calls
 *
 * @param headerRow  values of row 1
 * @returns column positions, null if any form column is missing
 */
function _matchFormHeader_(headerRow: unknown[]): FormColumns | null {
  const cells = headerRow.map((cell) => String(cell).trim());
  const keys = Object.keys(FORM_HEADERS) as (keyof typeof FORM_HEADERS)[];
  const columns: Partial<FormColumns> = {};
  for (const key of keys) {
    const index = cells.indexOf(FORM_HEADERS[key]);
    if (index === -1) return null;
    columns[key] = index;
  }
  return columns as FormColumns;
}

/**
 * find the responses tab and the Deleted forms tab of a Form Responses file, by their header rows
 * pure, no google calls
 *
 * a tab with every form header is the responses tab, today "Form responses 2"
 * one that also has "Reason for deletion" is the Deleted forms tab
 * there must be exactly one responses tab, the Deleted forms tab is optional
 *
 * @param tabs  every tab of the file
 * @returns the two tabs, or why they could not be found
 */
function _findFormTabs_(tabs: ResultsTab[]): FormTabsLookup {
  const responses: FormTab[] = [];
  const deleted: (FormTab & { reasonColumn: number })[] = [];
  for (const tab of tabs) {
    const header = tab.values[0] ?? [];
    const columns = _matchFormHeader_(header);
    if (!columns) continue;
    const reasonColumn = header.map((cell) => String(cell).trim()).indexOf(DELETED_REASON_HEADER);
    if (reasonColumn === -1) responses.push({ tab, columns });
    else deleted.push({ tab, columns, reasonColumn });
  }

  const names = (found: FormTab[]) => found.map((f) => `"${f.tab.title}"`).join(", ");
  if (responses.length === 0) return { ok: false, reason: "no tab has the form's headers" };
  if (responses.length > 1) {
    return { ok: false, reason: `${responses.length} tabs have the form's headers: ${names(responses)}` };
  }
  if (deleted.length > 1) {
    return { ok: false, reason: `${deleted.length} tabs have "${DELETED_REASON_HEADER}": ${names(deleted)}` };
  }
  return { ok: true, responses: responses[0], deleted: deleted[0] ?? null };
}

/**
 * the five category scores for one game, each a whole number 0–4
 */
type SpiritScores = Record<typeof SCORE_KEYS[number], number>;

/**
 * one usable row from the responses tab
 */
interface SpiritResponse {
  /** 1-based row number in the responses tab */
  sourceRow: number;
  /** when the form was submitted, null if blank */
  timestamp: Date | null;
  /** team giving the score ("Your Team Name") */
  scorer: string;
  /** team receiving the score ("Opponent Team Name") */
  receiver: string;
  /** the five category scores */
  scores: SpiritScores;
  /** spirit comment, "" if blank */
  comment: string;
}

/**
 * a row that could not be read, and why
 */
interface RowProblem {
  /** 1-based row number in the responses tab */
  sourceRow: number;
  /** what is wrong with the row */
  reason: string;
}

/**
 * one row of the Deleted forms tab, kept as it is: scores that are not 0–4 are null
 */
interface DeletedForm {
  /** 1-based row number in the Deleted forms tab */
  sourceRow: number;
  /** when the form was submitted, null if blank */
  timestamp: Date | null;
  /** team giving the score */
  scorer: string;
  /** team receiving the score */
  receiver: string;
  /** the five category scores, null where missing or not 0–4 */
  scores: Record<typeof SCORE_KEYS[number], number | null>;
  /** spirit comment, "" if blank */
  comment: string;
  /** "Reason for deletion" */
  reason: string;
}

/**
 * placeholder text some files have in the team column of a row that is not a response
 */
const PLACEHOLDER_TEAM = FORM_HEADERS.yourTeam;

/**
 * whether a cell holds a valid spirit score: a whole number from 0 to 4
 *
 * @param value cell value
 * @return true if valid
 */
function _isValidScore_(value: unknown): boolean {
  return typeof value === "number" && Number.isInteger(value) && value >= 0 && value <= 4;
}

/**
 * a timestamp cell as a date
 * pure, no google calls
 *
 * @param value  a date, or a Sheets serial number (days since 1899-12-30) as the Sheets API returns dates
 * @returns the date and time, null if the cell is not a date
 */
function _serialDateTime_(value: unknown): Date | null {
  if (value instanceof Date) return value;
  if (typeof value !== "number" || !Number.isFinite(value)) return null;
  const days = Math.floor(value);
  return new Date(1899, 11, 30 + days, 0, 0, 0, Math.round((value - days) * 86400000));
}

/**
 * turn the values of the responses tab into responses
 *
 * row 1 is the header and is skipped
 * completely blank rows and placeholder rows are skipped
 * rows with a missing team name or an invalid score are returned as problems instead of responses
 *
 * @param values   all values of the responses tab, including the header row
 * @param columns  column positions from _matchFormHeader_
 * @returns usable rows and unusable rows
 */
function _parseResponseRows_(
  values: unknown[][],
  columns: FormColumns,
): { responses: SpiritResponse[]; problems: RowProblem[] } {
  const responses: SpiritResponse[] = [];
  const problems: RowProblem[] = [];

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const sourceRow = i + 1;
    const scorer = _tidyTeamName_(row[columns.yourTeam]);
    const receiver = _tidyTeamName_(row[columns.opponentTeam]);
    const comment = String(row[columns.comments]).trim();

    const isBlank = !scorer && !receiver && !comment && SCORE_KEYS.every((k) => row[columns[k]] === "");
    if (isBlank || scorer === PLACEHOLDER_TEAM) continue;

    const reasons: string[] = [];
    if (!scorer) reasons.push(`missing ${FORM_HEADERS.yourTeam}`);
    if (!receiver) reasons.push(`missing ${FORM_HEADERS.opponentTeam}`);
    const badScores = SCORE_KEYS.filter((k) => !_isValidScore_(row[columns[k]]));
    if (badScores.length > 0) {
      reasons.push(
        `invalid score in ${
          badScores.map((k) => `"${FORM_HEADERS[k]}" (${JSON.stringify(row[columns[k]])})`).join(", ")
        }`,
      );
    }

    if (reasons.length > 0) {
      problems.push({ sourceRow: sourceRow, reason: reasons.join("; ") });
      continue;
    }

    responses.push({
      sourceRow: sourceRow,
      timestamp: _serialDateTime_(row[columns.timestamp]),
      scorer: scorer,
      receiver: receiver,
      scores: {
        rules: Number(row[columns.rules]),
        fouls: Number(row[columns.fouls]),
        fairMindedness: Number(row[columns.fairMindedness]),
        attitude: Number(row[columns.attitude]),
        communication: Number(row[columns.communication]),
      },
      comment: comment,
    });
  }
  return { responses: responses, problems: problems };
}

/**
 * turn the values of the Deleted forms tab into deleted forms, as they are
 * pure, no google calls
 *
 * row 1 is the header and is skipped, as are completely blank rows and placeholder rows
 * every other row is kept, it is a record of what the chair did, scores that are not 0–4 become null
 *
 * @param values        all values of the Deleted forms tab, including the header row
 * @param columns       column positions from _matchFormHeader_
 * @param reasonColumn  0-based position of "Reason for deletion"
 * @returns the deleted forms in tab order
 */
function _parseDeletedRows_(values: unknown[][], columns: FormColumns, reasonColumn: number): DeletedForm[] {
  const deleted: DeletedForm[] = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const scorer = _tidyTeamName_(row[columns.yourTeam]);
    const receiver = _tidyTeamName_(row[columns.opponentTeam]);
    const comment = String(row[columns.comments] ?? "").trim();
    const reason = String(row[reasonColumn] ?? "").trim();
    const isBlank = !scorer && !receiver && !comment && !reason && SCORE_KEYS.every((k) => row[columns[k]] === "");
    if (isBlank || scorer === PLACEHOLDER_TEAM) continue;

    const score = (key: typeof SCORE_KEYS[number]) => {
      const value = row[columns[key]];
      return _isValidScore_(value) ? value as number : null;
    };
    deleted.push({
      sourceRow: i + 1,
      timestamp: _serialDateTime_(row[columns.timestamp]),
      scorer,
      receiver,
      scores: {
        rules: score("rules"),
        fouls: score("fouls"),
        fairMindedness: score("fairMindedness"),
        attitude: score("attitude"),
        communication: score("communication"),
      },
      comment,
      reason,
    });
  }
  return deleted;
}

/**
 * tidy team name so the same team always has the same text
 * curly apostrophes -> straight
 * runs of whitespace -> single space
 * trimmed
 *
 * @param value  cell value
 * @returns the tidied name
 */
function _tidyTeamName_(value: unknown): string {
  return String(value ?? "").replace(/[‘’]/g, "'").replace(/\s+/g, " ").trim();
}
