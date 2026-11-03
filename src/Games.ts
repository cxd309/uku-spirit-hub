/**
 * name of the Games tab
 */
const GAMES_SHEET = "Games";

/**
 * text for the info row above the Games table
 * what the tab is, what to edit, how it refreshes
 */
const GAMES_INFO = "Every game on the Gamelist tab of each included tournament's schedule\n\n"
  + "DO NOT EDIT, this is all regenerated on refresh. Sort and filter is safe but will be overwritten on refresh\n\n"
  + "Each schedule is imported once, even when several divisions on the Tournaments tab share it\n"
  + "Games without both teams are left out. Played is unticked for 1-0 results, which are forfeits or cancelled games\n"
  + "Blank scores were not entered, the game still counts as played\n\n"
  + "To Refresh run \"Refresh Results\"";

/**
 * name of the tab in a schedule that lists every game
 */
const GAMELIST_SHEET = "Gamelist";

/**
 * 0-based Gamelist columns the hub reads, by position
 * the score columns have no header and both team columns are headed "Team", so they cannot be found by name
 */
const GAMELIST_COLUMNS = Object.freeze({
  game: 0,
  day: 2,
  time: 3,
  pitch: 4,
  team1: 5,
  score1: 6,
  score2: 8,
  team2: 9,
});

/**
 * games tab columns, in order
 * keys are the names used in code
 * values are the header text
 * columns listed in GAME_FORMULAS are formulas, the rest hold GameRecord values
 */
const GAME_HEADERS = Object.freeze(
  {
    scheduleUrl: "Schedule URL",
    eventName: "Schedule Event Name",
    sourceRow: "Source Row",
    game: "Game",
    day: "Day",
    time: "Time",
    pitch: "Pitch",
    team1: "Team 1",
    score1: "Score 1",
    score2: "Score 2",
    team2: "Team 2",
    played: "Played",
  } as const,
);

/**
 * keys of GameRecord in column order
 */
const GAME_KEYS = Object.keys(GAME_HEADERS) as (keyof typeof GAME_HEADERS)[];

/**
 * one row of the Games tab: one game in one schedule
 */
interface GameRecord {
  /** link to the schedule the game is from (links to the Tournaments tab) */
  scheduleUrl: string;
  /** row on the Gamelist tab */
  sourceRow: number;
  /** game code, e.g. A1vA3, 5v8, [2v3] */
  game: string;
  /** day of the game, null if blank */
  day: Date | null;
  /** start time as a fraction of a day, as Sheets stores times, null if blank */
  time: number | null;
  /** pitch, as written on the schedule */
  pitch: string;
  /** left team */
  team1: string;
  /** left team's score, null if not entered */
  score1: number | null;
  /** right team's score, null if not entered */
  score2: number | null;
  /** right team */
  team2: string;
  /** false for 1-0 results: forfeits and cancelled games */
  played: boolean;
}

/**
 * formula columns of the Games tab
 * each row gets its own formula, so sorting the tab never breaks them
 */
const GAME_FORMULAS: Readonly<Partial<Record<keyof typeof GAME_HEADERS, (row: number) => string>>> = Object.freeze({
  eventName: (row: number) => {
    const url = `$${_columnLetter_(GAME_KEYS.indexOf("scheduleUrl") + 1)}${row}`;
    const events = (key: keyof typeof EVENT_HEADERS) => _columnBelowHeader_(EVENTS_SHEET, EVENT_KEYS.indexOf(key) + 1);
    return `=IFERROR(XLOOKUP(${url}, ${events("scheduleUrl")}, ${events("scheduleEventName")}), "(unknown event)")`;
  },
});

/**
 * get the Games tab, creating it with header row on first use
 *
 * @returns the Games tab
 */
function _getGamesSheet_(): GoogleAppsScript.Spreadsheet.Sheet {
  const headers: readonly string[] = Object.values(GAME_HEADERS);
  return _getOrCreateSheet_(SpreadsheetApp.getActiveSpreadsheet(), GAMES_SHEET, headers, GAMES_INFO);
}

/**
 * a time as Sheets displays it, as a fraction of a day
 * pure, no google calls
 *
 * @param text  e.g. "09:30" or "9:30:00", as getDisplayValues gives it
 * @returns the fraction of a day, null if blank or not a time
 */
function _timeFromText_(text: string): number | null {
  const match = /^(\d{1,2}):(\d{2})(?::(\d{2}))?$/.exec(text.trim());
  if (!match) return null;
  return (Number(match[1]) * 3600 + Number(match[2]) * 60 + Number(match[3] ?? 0)) / 86400;
}

/**
 * @param row       values of one data row, in GAME_KEYS order
 * @param timeText  the Time cell as displayed: getValues gives times as dates in 1899, which do not read back reliably
 * @returns the record
 */
function _gameFromRow_(row: unknown[], timeText: string): GameRecord {
  const cell = (key: keyof typeof GAME_HEADERS) => row[GAME_KEYS.indexOf(key)];
  const number = (key: keyof typeof GAME_HEADERS) => (typeof cell(key) === "number" ? cell(key) as number : null);
  return {
    scheduleUrl: String(cell("scheduleUrl")),
    sourceRow: Number(cell("sourceRow")),
    game: String(cell("game")),
    day: cell("day") instanceof Date ? cell("day") as Date : null,
    time: _timeFromText_(timeText),
    pitch: String(cell("pitch")),
    team1: String(cell("team1")),
    score1: number("score1"),
    score2: number("score2"),
    team2: String(cell("team2")),
    played: cell("played") === true,
  };
}

/**
 * @param game  the record
 * @returns values in GAME_KEYS order, "" in formula columns (filled by _fillFormulaColumns_) and for blanks
 */
function _gameToRow_(game: GameRecord): unknown[] {
  return GAME_KEYS.map((key) => {
    if (GAME_FORMULAS[key]) return "";
    const value = game[key as keyof GameRecord];
    return value === null ? "" : value;
  });
}

/**
 * @param sheet  the Games tab
 * @returns every game on the tab
 */
function _readGames_(sheet: GoogleAppsScript.Spreadsheet.Sheet): GameRecord[] {
  const rowCount = sheet.getLastRow() - HEADER_ROW;
  if (rowCount < 1) return [];
  const times = sheet.getRange(DATA_ROW, GAME_KEYS.indexOf("time") + 1, rowCount, 1).getDisplayValues();
  return sheet
    .getRange(DATA_ROW, 1, rowCount, GAME_KEYS.length)
    .getValues()
    .map((row, i) => _gameFromRow_(row, times[i][0]))
    .filter((g) => g.scheduleUrl !== "");
}

/**
 * replace the contents of the Games tab, resizing it to fit
 * values are written first, then each formula column is filled in one call
 *
 * @param sheet  the Games tab
 * @param games  all games, in display order
 */
function _writeGames_(sheet: GoogleAppsScript.Spreadsheet.Sheet, games: GameRecord[]) {
  _writeTable_(sheet, games.map(_gameToRow_), GAME_KEYS.length);
  if (games.length === 0) return;
  _fillFormulaColumns_(sheet, GAME_KEYS, GAME_FORMULAS, games.length);
  const range = (key: keyof typeof GAME_HEADERS) =>
    sheet.getRange(DATA_ROW, GAME_KEYS.indexOf(key) + 1, games.length, 1);
  range("day").setNumberFormat("yyyy-mm-dd");
  range("time").setNumberFormat("hh:mm");
  range("played").setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
}

/**
 * games in tournament order, then as they are on the Gamelist
 * pure, returns a new array
 *
 * @param games   games in any order
 * @param events  events, used to find each schedule's date
 * @returns games in date order
 */
function _sortGames_(games: GameRecord[], events: EventRecord[]): GameRecord[] {
  const order = new Map<string, number>();
  _sortEvents_(events).forEach((e, i) => {
    if (!order.has(e.scheduleUrl)) order.set(e.scheduleUrl, i);
  });
  const position = (g: GameRecord) => order.get(g.scheduleUrl) ?? Number.MAX_SAFE_INTEGER;
  return [...games].sort((a, b) =>
    position(a) - position(b) || a.scheduleUrl.localeCompare(b.scheduleUrl) || a.sourceRow - b.sourceRow
  );
}

/**
 * the games of one schedule, from the values of its Gamelist tab
 * pure, no google calls
 *
 * row 1 is the header and is skipped, columns are read by position whatever it says
 * rows with no game code, or without both teams, are left out
 * error values such as #N/A count as blank
 *
 * @param scheduleUrl  link to the schedule, stored on each game
 * @param values       every value on the Gamelist from row 1, dates and times as serial numbers
 * @returns the games
 */
function _parseGamelist_(scheduleUrl: string, values: unknown[][]): GameRecord[] {
  const cell = (row: unknown[], key: keyof typeof GAMELIST_COLUMNS) => row[GAMELIST_COLUMNS[key]] ?? "";
  const text = (row: unknown[], key: keyof typeof GAMELIST_COLUMNS) => {
    const value = String(cell(row, key)).trim();
    return value.startsWith("#") ? "" : value;
  };
  const number = (row: unknown[], key: keyof typeof GAMELIST_COLUMNS) => {
    const value = cell(row, key);
    return typeof value === "number" && Number.isFinite(value) ? value : null;
  };

  const games: GameRecord[] = [];
  values.slice(1).forEach((row, i) => {
    const game = text(row, "game");
    const team1 = _tidyTeamName_(text(row, "team1"));
    const team2 = _tidyTeamName_(text(row, "team2"));
    if (game === "" || team1 === "" || team2 === "") return;

    const score1 = number(row, "score1");
    const score2 = number(row, "score2");
    const forfeit = (score1 === 1 && score2 === 0) || (score1 === 0 && score2 === 1);
    const time = number(row, "time");
    games.push({
      scheduleUrl,
      sourceRow: i + 2,
      game,
      day: _scheduleDate_(cell(row, "day")),
      time: time === null ? null : time - Math.floor(time),
      pitch: text(row, "pitch"),
      team1,
      score1,
      score2,
      team2,
      played: !forfeit,
    });
  });
  return games;
}

/**
 * a schedule's Gamelist values, or why they could not be read
 */
type GamelistRead = { ok: true; values: unknown[][] } | { ok: false; reason: string };

/**
 * fetch the Gamelist tab of many schedules at once
 * dates and times come back as serial numbers, as _parseGamelist_ expects
 *
 * @param urls  schedule links, each a Google Sheets link
 * @returns link → every value on its Gamelist from row 1, or why it could not be read
 */
function _readGamelists_(urls: string[]): Map<string, GamelistRead> {
  const responses = _fetchGoogleJson_(
    urls.map((url) =>
      `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(_spreadsheetIdFromUrl_(url))}/values/`
      + `${encodeURIComponent(`'${GAMELIST_SHEET}'`)}`
      + "?valueRenderOption=UNFORMATTED_VALUE&dateTimeRenderOption=SERIAL_NUMBER"
    ),
  );
  return new Map(urls.map((url, i): [string, GamelistRead] => {
    const response = responses[i];
    if (response.code === 200) {
      return [url, { ok: true, values: (response.body as GoogleAppsScript.Sheets.Schema.ValueRange).values ?? [] }];
    }
    if (response.code === 400) return [url, { ok: false, reason: `no "${GAMELIST_SHEET}" tab` }];
    if (response.code === 404) return [url, { ok: false, reason: "file not found, or not shared with you" }];
    return [url, { ok: false, reason: `could not open file (${_apiError_(response)})` }];
  }));
}
