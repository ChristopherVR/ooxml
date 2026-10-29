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
//     only through the trend, and the pair is kept when its trend is roughly 0.4 to 2 times the slope.
import { describe, expect, it } from 'vitest';
import { driftProbes } from './ets-drift-compare';
import { reportedFor, trendShape } from './functions/ets-search-map';
import { etsHoldout } from './functions/ets-core';

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
});
