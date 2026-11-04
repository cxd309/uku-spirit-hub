/**
 * values of the Match column on the Results tab
 */
const MATCH_RESULTS = Object.freeze(
  {
    OK: "OK",
    DUPLICATE: "Duplicate",
    SELF_SCORE: "Self score",
    NO_SUCH_GAME: "No such game",
    NO_SCHEDULE: "No schedule",
  } as const,
);

/**
 * whether a game's result is a 1-0, a forfeit or cancelled game
 * pure, no google calls
 *
 * @param game  the game
 * @returns true for 1-0 either way round
 */
function _isForfeit_(game: GameRecord): boolean {
  return (game.score1 === 1 && game.score2 === 0) || (game.score1 === 0 && game.score2 === 1);
}

/**
 * match every response to a game, and mark each game with which teams submitted for it
 * pure, never modifies its arguments
 *
 * only tournaments with an imported schedule that is not ERROR or MISSING are matched,
 * the rest are No schedule and count as they are
 * for each scoring team → receiving team at a tournament:
 * - a team scoring itself is a Self score
 * - with no game between them every form is No such game
 * - otherwise as many forms count as there were games, those with a comment first, then the latest,
 *   the counted forms go to the games in timestamp order and the rest are Duplicates
 * Self scores, No such game and Duplicates do not count
 *
 * a 1-0 game is not played, unless a form was matched to it
 *
 * @param responses  every response on the Results tab
 * @param games      every game on the Games tab
 * @param events     every event, to find each response's schedule
 * @returns the responses with Game, Counted and Match set, and the games with Played and Submitted set
 */
function _matchResponses_(
  responses: ResponseRecord[],
  games: GameRecord[],
  events: EventRecord[],
): { responses: ResponseRecord[]; games: GameRecord[] } {
  const usable = (e: EventRecord) =>
    e.scheduleImportDate !== null
    && e.scheduleStatus !== SOURCE_STATUS.ERROR
    && e.scheduleStatus !== SOURCE_STATUS.MISSING;
  const scheduleOf = new Map(events.filter(usable).map((e) => [e.spiritFormUrl, e.scheduleUrl]));
  const pairKey = (schedule: string, scorer: string, receiver: string) =>
    `${schedule}|${_teamKey_(scorer)}|${_teamKey_(receiver)}`;
  const gameKey = (g: GameRecord) => `${g.scheduleUrl}|${g.sourceRow}`;

  // games between each pair of teams, both ways round, earliest first
  const gamesByPair: Map<string, GameRecord[]> = new Map();
  const ordered = [...games].sort((a, b) =>
    _compareDates_(a.day, b.day) || (a.time ?? 0) - (b.time ?? 0) || a.sourceRow - b.sourceRow
  );
  for (const game of ordered) {
    for (const [scorer, receiver] of [[game.team1, game.team2], [game.team2, game.team1]]) {
      const key = pairKey(game.scheduleUrl, scorer, receiver);
      gamesByPair.set(key, [...(gamesByPair.get(key) ?? []), game]);
    }
  }

  const matched: Map<ResponseRecord, ResponseRecord> = new Map();
  const formsByPair: Map<string, ResponseRecord[]> = new Map();
  for (const r of responses) {
    const schedule = scheduleOf.get(r.spiritFormUrl);
    if (schedule === undefined) {
      matched.set(r, { ...r, game: "", counted: true, match: MATCH_RESULTS.NO_SCHEDULE });
    } else if (_teamKey_(r.scorer) === _teamKey_(r.receiver)) {
      matched.set(r, { ...r, game: "", counted: false, match: MATCH_RESULTS.SELF_SCORE });
    } else {
      const key = pairKey(schedule, r.scorer, r.receiver);
      formsByPair.set(key, [...(formsByPair.get(key) ?? []), r]);
    }
  }

  const submitted: Set<string> = new Set();
  const time = (r: ResponseRecord) => r.timestamp?.getTime() ?? 0;
  for (const [key, forms] of formsByPair) {
    const pairGames = gamesByPair.get(key) ?? [];
    const byPriority = [...forms].sort((a, b) =>
      Number(b.comment !== "") - Number(a.comment !== "") || time(b) - time(a) || b.sourceRow - a.sourceRow
    );
    const counted = byPriority.slice(0, pairGames.length).sort((a, b) =>
      time(a) - time(b) || a.sourceRow - b.sourceRow
    );
    counted.forEach((r, i) => {
      const game = pairGames[i];
      matched.set(r, { ...r, game: game.game, counted: true, match: MATCH_RESULTS.OK });
      submitted.add(`${gameKey(game)}|${_teamKey_(r.scorer)}`);
    });
    for (const r of byPriority.slice(pairGames.length)) {
      const match = pairGames.length === 0 ? MATCH_RESULTS.NO_SUCH_GAME : MATCH_RESULTS.DUPLICATE;
      matched.set(r, { ...r, game: "", counted: false, match });
    }
  }

  return {
    responses: responses.map((r) => matched.get(r) ?? r),
    games: games.map((g) => {
      const team1Submitted = submitted.has(`${gameKey(g)}|${_teamKey_(g.team1)}`);
      const team2Submitted = submitted.has(`${gameKey(g)}|${_teamKey_(g.team2)}`);
      return { ...g, played: !_isForfeit_(g) || team1Submitted || team2Submitted, team1Submitted, team2Submitted };
    }),
  };
}
