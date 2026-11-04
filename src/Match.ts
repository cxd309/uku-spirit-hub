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
 * whether a Deleted forms reason says the form was a duplicate
 * a duplicate is an extra form, not the form for a game, so it takes no place when forms are lined up with games
 *
 * @param reason  the "Reason for deletion" cell
 * @returns true when it mentions "duplicate", ignoring case
 */
function _isDuplicateReason_(reason: string): boolean {
  return /duplicate/i.test(reason);
}

/**
 * match every response and deleted form to a game, and mark each game with which teams submitted for it
 * pure, never modifies its arguments
 *
 * only tournaments with an imported schedule that is not ERROR or MISSING are matched,
 * the rest are No schedule and count as they are
 * for each scoring team → receiving team at a tournament:
 * - a team scoring itself is a Self score
 * - with no game between them every form is No such game
 * - otherwise the forms on Results and the deleted forms are taken in the order they were submitted,
 *   leaving out deleted copies of forms still on Results (an edited comment, same submission time) and duplicates,
 *   and each takes the earliest game without a form that had started when it was submitted
 * - with no game free, a form on Results takes the game of a deleted form (a resubmission),
 *   or of a form on Results it beats (a comment first, then the later one), which becomes a Duplicate
 * - a form on Results submitted before any of their games still takes a free game, a deleted form gets none
 * - forms on Results left without a game are Duplicates, deleted forms never count
 * Self scores, No such game and Duplicates do not count
 * a copy of a form still on Results takes that form's game
 *
 * a 1-0 game is not played, unless a form was matched to it
 *
 * @param responses  every response on the Results tab
 * @param deleted    every form on the Deleted Forms tab
 * @param games      every game on the Games tab
 * @param events     every event, to find each response's schedule
 * @returns the responses with Game, Counted and Match set, the deleted forms with Game set,
 *   and the games with Played and Submitted set
 */
function _matchResponses_(
  responses: ResponseRecord[],
  deleted: DeletedFormRecord[],
  games: GameRecord[],
  events: EventRecord[],
): { responses: ResponseRecord[]; deleted: DeletedFormRecord[]; games: GameRecord[] } {
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

  // deleted forms that stand for a game's form, and copies of forms still on Results
  const sameForm = (f: { spiritFormUrl: string; scorer: string; receiver: string; timestamp: Date | null }) =>
    `${f.spiritFormUrl}|${_teamKey_(f.scorer)}|${_teamKey_(f.receiver)}|${f.timestamp?.getTime() ?? ""}`;
  const keptForms = new Map(responses.filter((r) => r.timestamp !== null).map((r) => [sameForm(r), r]));
  const copies: Map<DeletedFormRecord, ResponseRecord> = new Map();
  const deletedByPair: Map<string, DeletedFormRecord[]> = new Map();
  for (const f of deleted) {
    const schedule = scheduleOf.get(f.spiritFormUrl);
    const copyOf = f.timestamp === null ? undefined : keptForms.get(sameForm(f));
    if (copyOf) copies.set(f, copyOf);
    else if (
      schedule !== undefined && !_isDuplicateReason_(f.reason) && _teamKey_(f.scorer) !== _teamKey_(f.receiver)
    ) {
      const key = pairKey(schedule, f.scorer, f.receiver);
      deletedByPair.set(key, [...(deletedByPair.get(key) ?? []), f]);
    }
  }

  const submitted: Set<string> = new Set();
  const deletedGame: Map<DeletedFormRecord, string> = new Map();
  const time = (f: { timestamp: Date | null }) => f.timestamp?.getTime() ?? 0;
  const startOf = (g: GameRecord) =>
    g.day === null
      ? null
      : new Date(g.day.getFullYear(), g.day.getMonth(), g.day.getDate(), 0, 0, 0, Math.round((g.time ?? 0) * 86400000))
        .getTime();
  // whether form a counts over form b for the same game: a comment first, then the later one
  const beats = (a: ResponseRecord, b: ResponseRecord) =>
    (Number(a.comment !== "") - Number(b.comment !== "") || time(a) - time(b) || a.sourceRow - b.sourceRow) > 0;

  for (const key of new Set([...formsByPair.keys(), ...deletedByPair.keys()])) {
    const pairGames = gamesByPair.get(key) ?? [];
    type Held = { kept: ResponseRecord; timestamp: Date | null; sourceRow: number } | {
      deleted: DeletedFormRecord;
      timestamp: Date | null;
      sourceRow: number;
    };
    const holders: (Held | null)[] = pairGames.map(() => null);
    const duplicates: ResponseRecord[] = [];
    const lineUp: Held[] = [
      ...(formsByPair.get(key) ?? []).map((r) => ({ kept: r, timestamp: r.timestamp, sourceRow: r.sourceRow })),
      ...(deletedByPair.get(key) ?? []).map((f) => ({ deleted: f, timestamp: f.timestamp, sourceRow: f.sourceRow })),
    ].sort((a, b) => time(a) - time(b) || a.sourceRow - b.sourceRow);

    for (const form of lineUp) {
      // a form can only be for a game that had started when it was submitted
      const started = pairGames.flatMap((g, i) => {
        const start = startOf(g);
        return form.timestamp === null || start === null || start <= form.timestamp.getTime() ? [i] : [];
      });
      // a form on Results submitted before any of their games is still given one, a deleted form is not
      const eligible = "kept" in form && started.length === 0 ? pairGames.map((_, i) => i) : started;

      const free = eligible.find((i) => holders[i] === null);
      if (free !== undefined) {
        holders[free] = form;
        continue;
      }
      if (!("kept" in form)) continue;
      // a resubmission takes the game of a deleted form, then a form it beats takes second place
      const fromDeleted = eligible.find((i) => holders[i] !== null && "deleted" in (holders[i] as Held));
      if (fromDeleted !== undefined) {
        holders[fromDeleted] = form;
        continue;
      }
      const weaker = eligible.find((i) => {
        const held = holders[i];
        return held !== null && "kept" in held && beats(form.kept, held.kept);
      });
      if (weaker === undefined) {
        duplicates.push(form.kept);
        continue;
      }
      const displaced = holders[weaker];
      if (displaced !== null && "kept" in displaced) duplicates.push(displaced.kept);
      holders[weaker] = form;
    }

    holders.forEach((held, i) => {
      if (held === null) return;
      const game = pairGames[i];
      if ("kept" in held) {
        matched.set(held.kept, { ...held.kept, game: game.game, counted: true, match: MATCH_RESULTS.OK });
        submitted.add(`${gameKey(game)}|${_teamKey_(held.kept.scorer)}`);
      } else {
        deletedGame.set(held.deleted, game.game);
        submitted.add(`${gameKey(game)}|${_teamKey_(held.deleted.scorer)}`);
      }
    });
    for (const r of duplicates) {
      const match = pairGames.length === 0 ? MATCH_RESULTS.NO_SUCH_GAME : MATCH_RESULTS.DUPLICATE;
      matched.set(r, { ...r, game: "", counted: false, match });
    }
  }

  const matchedResponses = responses.map((r) => matched.get(r) ?? r);
  const gameOfResponse = new Map(responses.map((r, i) => [r, matchedResponses[i].game]));
  return {
    responses: matchedResponses,
    deleted: deleted.map((f) => {
      const copyOf = copies.get(f);
      return { ...f, game: copyOf ? gameOfResponse.get(copyOf) ?? "" : deletedGame.get(f) ?? "" };
    }),
    games: games.map((g) => {
      const team1Submitted = submitted.has(`${gameKey(g)}|${_teamKey_(g.team1)}`);
      const team2Submitted = submitted.has(`${gameKey(g)}|${_teamKey_(g.team2)}`);
      return { ...g, played: !_isForfeit_(g) || team1Submitted || team2Submitted, team1Submitted, team2Submitted };
    }),
  };
}
