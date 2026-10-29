// Which forecast Excel makes when the lattice pair has a large beta: the pair's, or the last value plus
// the least-squares slope ("drift"). Labelled on the 742 recorded probes of 8 or more points whose
// lowest-cost pair has a beta above 0.001 (ets-drift-compare.ts).
//
// Measured:
//   - the drift forecast is exact: the filter with alpha 1 and beta 0 from the least-squares slope
//     (last value + h * slope) equals Excel's forecast to 1e-15 on all 440 drifting probes, and the
//     other 302 probes equal the pair's forecast; neither (alpha, 0.001) nor (0.998, 0.001) does
//     (3e-4 off), so it is a different branch and not a fit with a small beta;
//   - STAT 2 is 0.001 on exactly the drifting probes, the pair's beta on the others;
//   - the choice is not a function of one statistic: the best single threshold gets 74% (the slope
//     in residual sd), a depth-3 tree on shift, the Durbin-Watson of the linear fit and the pair's
//     trend over the slope 87% (85% cross-validated), the rule in ets-search-map.ts 83%. Along 9
//     blends of a drifting and a pair series the flip happens where no such statistic crosses a
//     common value (the Durbin-Watson comes closest). Where the alpha is 0.998 the two forecasts differ
//     only through the trend, and the pair is kept when its trend is roughly 0.4 to 2 times the slope;
//   - the choice is not a comparison of the two forecasts' errors either. Drift (alpha 1, beta 0, the
//     full-data slope, or the slope of the training part only) against the chosen pair, on the cost
//     the search uses, on its hold-out part alone, or as a naive hold-out forecast from the end of the
//     training part: the best single threshold on the cost ratio gets 442 to 485 of 742 (59% of them
//     drift), and even a threshold fitted separately for each of the 25 pairs (direction free) gets
//     564 to 590, under the 614 of the rule in ets-search-map.ts.
import { describe, expect, it } from 'vitest';
import { driftProbes } from './ets-drift-compare';
import { reportedFor, trendShape } from './functions/ets-search-map';
import { etsHoldout, etsRun } from './functions/ets-core';
import { initialStates } from './functions/ets-init';
import { etsSearchSeries, etsValidationCost } from './functions/ets-validate';

const probes = driftProbes().filter((p) => p.pair[1] !== 0.001);

describe('drift versus pair forecasts of FORECAST.ETS', () => {
	it('labels every probe as exactly one of the two forecasts', () => {
		expect(probes.length).toBe(742);
		expect(probes.filter((p) => p.drift === null)).toEqual([]);
		expect(probes.filter((p) => p.drift === true).length).toBe(440);
	});

	it('reports beta 0.001 exactly where it forecasts the drift', () => {
		for (const p of probes) expect(p.excelBeta === 0.001).toBe(p.drift === true);
	});

	it('predicts the choice on at least 614 of the 742 probes', () => {
		const right = probes.filter((p) => {
			const shape = trendShape(p.y, etsHoldout(p.y.length, 0));
			return reportedFor(p.pair[0], p.pair[1], shape).drift === p.drift;
		});
		expect(right.length).toBeGreaterThanOrEqual(614);
	});

	it('is not decided by the cost of the drift model against the pair, even per pair', () => {
		const ratio = (p: (typeof probes)[number]): number => {
			const y = p.y;
			const h = etsHoldout(y.length, 0);
			const searched = etsSearchSeries(y);
			const cost = (alpha: number, beta: number): number => {
				const q = { alpha, beta, gamma: 0.001 };
				return etsValidationCost(searched, 0, initialStates(y, 0, q, etsRun), q, h);
			};
			return cost(1, 0) / cost(p.pair[0], p.pair[1]);
		};
		const best = (items: { x: number; y: boolean }[]): number => {
			let top = 0;
			for (const t of items.map((i) => i.x))
				for (const above of [true, false])
					top = Math.max(top, items.filter((i) => (i.x >= t === above) === i.y).length);
			return top;
		};
		const items = probes.map((p) => ({ x: ratio(p), y: p.drift === true, key: p.pair.join('/') }));
		const perPair = [...new Set(items.map((i) => i.key))].reduce(
			(sum, key) => sum + best(items.filter((i) => i.key === key)),
			0,
		);
		expect(best(items)).toBeLessThan(500);
		expect(perPair).toBeLessThan(614);
	});
});
