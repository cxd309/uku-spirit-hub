/**
 * the spirit score model, the same as the legacy R script
 *   lmer(total ~ club + (1 | scorer / tournament))
 * a mean for each receiving club, adjusted for scorers who give higher or lower scores than others
 *   scorer: a random effect for each scorer
 *   scorer at tournament: a random effect for each scorer at each tournament, inside the scorer one
 * fitted by REML, like lme4, and checked against a reference implementation
 *
 * scorer at tournament always sits inside scorer, so each scorer's part of the fit
 * is a few sums with no big matrices, leaving one small club by club system
 */

/**
 * one score for the model
 *
 * @typedef {Object} ModelObservation
 * @property {string} club        receiving club
 * @property {string} scorer      scorer group, a club or a team
 * @property {string} tournament  tournament id
 * @property {number} total       total score
 */

/**
 * a fitted model, or why it could not be fitted
 *
 * @typedef {{
 *   ok: true,
 *   clubs: Map<string, {mean: number, se: number}>,
 *   scorerEffects: Map<string, number>,
 *   residualVariance: number,
 *   scorerVariance: number,
 *   scorerAtTournamentVariance: number
 * } | {
 *   ok: false,
 *   reason: string
 * }} ModelResult
 */

/**
 * the scores one scorer gave at one tournament
 *
 * @typedef {Object} ModelGroup
 * @property {number} count           how many scores
 * @property {number} sum             sum of the totals
 * @property {number} sumSquares      sum of the squared totals
 * @property {Map<number, {count: number, sum: number}>} byClub  club index -> scores given to that club
 */

/**
 * group the observations by scorer, then by tournament
 * pure, never modifies its arguments
 *
 * @param {ModelObservation[]}  observations  the scores
 * @param {Map<string, number>} clubIndex     club → column
 * @returns {Map<string, ModelGroup[]>} scorer → one group per tournament
 */
function _modelGroups_(observations, clubIndex) {
  /** @type {Map<string, Map<string, ModelGroup>>} */
  const byScorer = new Map();
  for (const o of observations) {
    const tournaments = byScorer.get(o.scorer) ?? new Map();
    byScorer.set(o.scorer, tournaments);
    const group = tournaments.get(o.tournament) ?? { count: 0, sum: 0, sumSquares: 0, byClub: new Map() };
    tournaments.set(o.tournament, group);
    const j = /** @type {number} */ (clubIndex.get(o.club));
    const cell = group.byClub.get(j) ?? { count: 0, sum: 0 };
    group.byClub.set(j, { count: cell.count + 1, sum: cell.sum + o.total });
    group.count++;
    group.sum += o.total;
    group.sumSquares += o.total * o.total;
  }
  return new Map([...byScorer].map(([scorer, tournaments]) => [scorer, [...tournaments.values()]]));
}

/**
 * cholesky factor of a symmetric matrix
 *
 * @param {Float64Array[]} m  the matrix
 * @returns {Float64Array[]|null} lower triangle L with L Lᵀ = m, null if m is not positive definite
 */
function _cholesky_(m) {
  const size = m.length;
  const l = Array.from({ length: size }, () => new Float64Array(size));
  for (let i = 0; i < size; i++) {
    for (let j = 0; j <= i; j++) {
      let sum = m[i][j];
      for (let k = 0; k < j; k++) sum -= l[i][k] * l[j][k];
      if (i === j) {
        if (!(sum > 0)) return null;
        l[i][i] = Math.sqrt(sum);
      } else {
        l[i][j] = sum / l[j][j];
      }
    }
  }
  return l;
}

/**
 * solve L Lᵀ x = v
 *
 * @param {Float64Array[]}         l  cholesky factor
 * @param {ArrayLike<number>}      v  right hand side
 * @returns {Float64Array} x
 */
function _choleskySolve_(l, v) {
  const size = l.length;
  const z = new Float64Array(size);
  for (let i = 0; i < size; i++) {
    let sum = v[i];
    for (let k = 0; k < i; k++) sum -= l[i][k] * z[k];
    z[i] = sum / l[i][i];
  }
  const x = new Float64Array(size);
  for (let i = size - 1; i >= 0; i--) {
    let sum = z[i];
    for (let k = i + 1; k < size; k++) sum -= l[k][i] * x[k];
    x[i] = sum / l[i][i];
  }
  return x;
}

/**
 * minimise a function of two numbers with the nelder-mead method
 *
 * @param {function([number, number]): number} f      function to minimise
 * @param {[number, number]}                   start  starting point
 * @returns {[number, number]} the minimum found
 */
function _nelderMead2_(f, start) {
  /** @param {[number, number]} x */
  const point = (x) => ({ x, value: f(x) });
  let simplex = [start, [start[0] + 0.5, start[1]], [start[0], start[1] + 0.5]].map(
    (x) => point(/** @type {[number, number]} */ (x)),
  );
  for (let i = 0; i < 2000; i++) {
    simplex.sort((a, b) => a.value - b.value);
    const [best, middle, worst] = simplex;
    const size = Math.max(...simplex.map((p) => Math.abs(p.x[0] - best.x[0]) + Math.abs(p.x[1] - best.x[1])));
    if (worst.value - best.value < 1e-10 && size < 1e-8) break;

    const centre = [(best.x[0] + middle.x[0]) / 2, (best.x[1] + middle.x[1]) / 2];
    /** @param {number} k */
    const along = (k) => point([centre[0] + k * (worst.x[0] - centre[0]), centre[1] + k * (worst.x[1] - centre[1])]);
    const reflected = along(-1);
    if (reflected.value < best.value) {
      const expanded = along(-2);
      simplex[2] = expanded.value < reflected.value ? expanded : reflected;
    } else if (reflected.value < middle.value) {
      simplex[2] = reflected;
    } else {
      const contracted = along(reflected.value < worst.value ? -0.5 : 0.5);
      if (contracted.value < Math.min(reflected.value, worst.value)) {
        simplex[2] = contracted;
      } else {
        simplex = [best, ...[middle, worst].map((p) => point([(p.x[0] + best.x[0]) / 2, (p.x[1] + best.x[1]) / 2]))];
      }
    }
  }
  simplex.sort((a, b) => a.value - b.value);
  return simplex[0].x;
}

/**
 * fit the model
 * pure, no google calls
 *
 * each scorer's scores are correlated through the two random effects
 * with variance ratios t1 (scorer) and t2 (scorer at tournament) the inverse of that correlation is exact:
 *   within a tournament group of n scores, weight w = 1 / (1 + t2 n)
 *   across the scorer, s = sum of n w over its groups, and the scorer part shrinks by t1 / (1 + t1 s)
 * the club means are then generalised least squares, and REML picks t1 and t2
 *
 * @param {ModelObservation[]} observations  the scores
 * @returns {ModelResult} club means with standard errors, or why the fit failed
 */
function _fitClubModel_(observations) {
  const clubs = [...new Set(observations.map((o) => o.club))].sort();
  const clubIndex = new Map(clubs.map((club, i) => [club, i]));
  const p = clubs.length;
  const n = observations.length;
  if (n <= p) return { ok: false, reason: `not enough scores, ${n} scores for ${p} clubs` };
  const groups = _modelGroups_(observations, clubIndex);

  /**
   * the fit for given relative standard deviations, squared to give the variance ratios
   * @param {[number, number]} theta
   */
  const fit = (theta) => {
    const t1 = theta[0] ** 2;
    const t2 = theta[1] ** 2;
    const xvx = Array.from({ length: p }, () => new Float64Array(p));
    const xvy = new Float64Array(p);
    let yvy = 0;
    let logDetV = 0;
    for (const tournaments of groups.values()) {
      let s = 0;
      let ay = 0;
      /** @type {Map<number, number>} */
      const a = new Map();
      for (const g of tournaments) {
        const w = 1 / (1 + t2 * g.count);
        logDetV += Math.log(1 + t2 * g.count);
        s += g.count * w;
        ay += w * g.sum;
        yvy += g.sumSquares - t2 * w * g.sum * g.sum;
        for (const [j, cell] of g.byClub) {
          a.set(j, (a.get(j) ?? 0) + w * cell.count);
          xvy[j] += cell.sum - t2 * w * cell.count * g.sum;
          xvx[j][j] += cell.count;
          for (const [k, other] of g.byClub) xvx[j][k] -= t2 * w * cell.count * other.count;
        }
      }
      const shrink = t1 / (1 + t1 * s);
      logDetV += Math.log(1 + t1 * s);
      yvy -= shrink * ay * ay;
      for (const [j, aj] of a) {
        xvy[j] -= shrink * aj * ay;
        for (const [k, ak] of a) xvx[j][k] -= shrink * aj * ak;
      }
    }
    const l = _cholesky_(xvx);
    if (!l) return null;
    const beta = _choleskySolve_(l, xvy);
    const rss = yvy - beta.reduce((sum, b, j) => sum + b * xvy[j], 0);
    const sigma2 = rss / (n - p);
    const logDetXvx = l.reduce((sum, row, i) => sum + 2 * Math.log(row[i]), 0);
    const deviance = logDetV + logDetXvx + (n - p) * (Math.log(2 * Math.PI * sigma2) + 1);
    return { deviance, beta, sigma2, l, t1, t2 };
  };

  const best = fit(_nelderMead2_((theta) => fit(theta)?.deviance ?? Infinity, [0.5, 0.5]));
  if (!best || !(best.sigma2 > 0)) return { ok: false, reason: "the model could not be fitted to these scores" };

  const clubResults = new Map(clubs.map((club, j) => {
    const unit = new Float64Array(p);
    unit[j] = 1;
    return [club, { mean: best.beta[j], se: Math.sqrt(best.sigma2 * _choleskySolve_(best.l, unit)[j]) }];
  }));

  // each scorer's predicted effect: t1 × (weighted residual sum) / (1 + t1 s)
  const scorerEffects = new Map([...groups].map(([scorer, tournaments]) => {
    let s = 0;
    let residual = 0;
    for (const g of tournaments) {
      const w = 1 / (1 + best.t2 * g.count);
      const fitted = [...g.byClub].reduce((sum, [j, cell]) => sum + cell.count * best.beta[j], 0);
      s += g.count * w;
      residual += w * (g.sum - fitted);
    }
    return [scorer, (best.t1 * residual) / (1 + best.t1 * s)];
  }));

  return {
    ok: true,
    clubs: clubResults,
    scorerEffects,
    residualVariance: best.sigma2,
    scorerVariance: best.t1 * best.sigma2,
    scorerAtTournamentVariance: best.t2 * best.sigma2,
  };
}
