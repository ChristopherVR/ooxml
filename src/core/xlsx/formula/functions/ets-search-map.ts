// The (alpha, beta) lattice of Excel's FORECAST.ETS search without a season, and what Excel reports
// for each optimum on it. Found from about 1500 probe series (ets-search.ts has the evidence).
//
// The candidates are the pairs Excel is seen to fit with (the error statistics of a series
// reproduce to the last digit with exactly these pairs): beta 0.001 with eight alphas, then
// alpha - 0.001 and a few larger betas. The lowest hold-out cost over them is Excel's pair for 99.6%
// of the series that the statistics identify. Nearby pairs (the `PROXIES`) are not Excel's own, but
// when one of them is the optimum Excel's answer is one fixed outcome, which removes most of the
// remaining differences.
//
// What Excel reports from the optimum is not always the pair itself. For a large beta (alpha, beta
// next to alpha - 0.001, or 0.249, 0.499 with alpha 0.998 half of the time) STAT 2 says 0.001
// and the forecast is the last value plus the least-squares slope (alpha 1, beta 0), while the
// error statistics still use the pair. Each entry below is the commonest outcome of the probes with
// that optimum.

export interface Outcome {
	/** The alpha FORECAST.ETS.STAT reports. */
	alpha: number;
	/** The beta it reports. */
	beta: number;
	/** The pair the error statistics are measured with. */
	fit: readonly [number, number];
	/** Excel forecasts from the last value plus the least-squares slope. */
	drift: boolean;
}

const LOW = 0.001;
/** Relative cost added to alpha 0.998 with the lowest beta (Excel avoids the upper end). */
export const TOP_COST = 0.05;

/** Candidates in the order they are tried (a later one wins only by a strictly lower cost). */
export const CANDIDATES: readonly (readonly [number, number])[] = [
	...[0.002, 0.1, 0.126, 0.25, 0.5, 0.75, 0.9, 0.998].map((a) => [a, LOW] as const),
	[0.1, 0.099],
	[0.126, 0.125],
	[0.25, 0.249],
	[0.333333, 0.166667],
	[0.5, 0.25],
	[0.5, 0.499],
	[0.666667, 0.166667],
	[0.666667, 0.5],
	[0.75, 0.749],
	[0.833333, 0.75],
	[0.9, 0.899],
	[0.998, 0.099],
	[0.998, 0.249],
	[0.998, 0.499],
	[0.998, 0.749],
	[0.998, 0.899],
	[0.166667, 0.0833333],
	// The proxies: pairs that are not Excel's own but stand for one outcome each.
	[0.9, 0.099],
	[0.9, 0.499],
	[0.9, 0.249],
	[0.998, 0.125],
	[0.75, 0.125],
];

const rep = (alpha: number, beta: number = LOW): Outcome => ({
	alpha,
	beta,
	fit: [alpha, beta],
	drift: false,
});
/** Reported alpha with beta 0.001, forecast from the last value; the statistics use `fitBeta`. */
const drift = (alpha: number, fitBeta: number = LOW): Outcome => ({
	alpha,
	beta: LOW,
	fit: [alpha, fitBeta],
	drift: true,
});

/** Optimum `alpha/beta` to what Excel reports (others are reported as they are). */
const OUTCOMES: Record<string, Outcome> = {
	'0.1/0.099': drift(0.1, 0.099), // 10 of 11
	'0.126/0.125': rep(0.126, 0.125),
	'0.166667/0.0833333': drift(0.166667, 0.0833333), // 35 of 36; the pair itself, found from the error statistics
	'0.25/0.249': drift(0.25, 0.249), // 54 of 55
	'0.333333/0.166667': rep(0.333333, 0.166667),
	'0.5/0.25': drift(0.5, 0.25), // 8 of 11
	'0.5/0.499': drift(0.5, 0.499), // 14 of 15
	'0.666667/0.166667': rep(0.666667, 0.166667), // 14 of 25, the rest forecast the drift
	'0.666667/0.5': drift(0.666667, 0.5), // 7 of 8
	'0.75/0.125': drift(0.833333, 0.0833333), // 2 of 3
	'0.833333/0.75': drift(0.833333, 0.75), // 2 of 2
	'0.9/0.099': rep(0.998, 0.099), // 16 of 16
	'0.9/0.249': drift(0.998, 0.249), // 11 of 14
	'0.9/0.499': drift(0.998, 0.499), // 18 of 23
	'0.9/0.899': drift(0.9, 0.899), // 12 of 19, 7 report 1/6
	'0.998/0.001': rep(0.998, 0.099), // 3 of 5, 2 report 0.9 (the 5% above keeps this rare)
	'0.998/0.099': rep(0.998, 0.099),
	'0.998/0.125': drift(0.998, 0.099), // 18 of 20
	'0.998/0.249': rep(0.998, 0.249), // 25 of 28
	'0.998/0.499': rep(0.998, 0.499),
	'0.998/0.749': drift(0.998, 0.749),
	'0.998/0.899': drift(0.998, 0.899),
};

/** What Excel reports when the lowest cost of the lattice is at (alpha, beta). */
export function reportedFor(alpha: number, beta: number, n: number): Outcome {
	// Three series have this optimum: 13 and 16 points report the pair, 24 points forecast the drift.
	if (alpha === 0.75 && beta === 0.749) return n >= 20 ? drift(0.75, 0.749) : rep(0.75, 0.749);
	return OUTCOMES[`${alpha}/${beta}`] ?? rep(alpha, beta);
}
