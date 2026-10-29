// The error Excel's FORECAST.ETS search minimises (identified from 1100 probe series, see
// ets-search.ts and docs/xlsx-parity-status.md): the series is split into a training part and the
// hold-out of `etsHoldout` trailing points. The filter runs over the training part one step at a
// time, then forecasts the hold-out several steps ahead without updating. The cost is the sum of
// the squared one-step errors of the training part plus the squared multi-step forecast errors of
// the hold-out. Series too short for a hold-out use the plain one-step error, and a series of 4 to 6
// points is first extended to 7 (`etsSearchSeries`).
import { type EtsInit, olsSlope } from './ets-init';

interface Params {
	alpha: number;
	beta: number;
	gamma: number;
}

/**
 * The points appended to a series of 4 to 6 values, as `[source, lag]` pairs: the point is
 * `y[source] + lag * slope`. The same extension gives the starting trend (`PADS` in `ets-init.ts`).
 */
const EXTENSION: Record<number, readonly (readonly [number, number])[]> = {
	6: [[0, 6]],
	5: [
		[1, 4],
		[0, 6],
	],
	4: [
		[0, 4],
		[1, 4],
		[0, 6],
	],
};

/**
 * The series Excel's search measures the one-step error over: 4 to 6 values are extended to 7 with
 * the same points the starting trend uses, and the search then picks Excel's parameters on all of
 * the 40 recorded series of 4 to 7 values (on the values alone it picks 29).
 */
export function etsSearchSeries(y: readonly number[]): readonly number[] {
	const pads = EXTENSION[y.length];
	if (!pads) return y;
	const slope = olsSlope(y);
	return [...y, ...pads.map(([src, lag]) => (y[src] as number) + lag * slope)];
}

/** The training-plus-hold-out squared error of the additive filter for these parameters. */
export function etsValidationCost(
	y: readonly number[],
	m: number,
	init: EtsInit,
	p: Params,
	holdout: number,
): number {
	const n = y.length;
	const train = n - holdout;
	let { level, trend } = init;
	const season = [...init.season];
	let sse = 0;
	for (let t = 0; t < train; t++) {
		const k = m >= 2 ? t % m : 0;
		const e = (y[t] as number) - (level + trend + (m >= 2 ? (season[k] ?? 0) : 0));
		sse += e * e;
		level += trend + p.alpha * e;
		trend += p.beta * e;
		if (m >= 2) season[k] = (season[k] ?? 0) + p.gamma * e;
	}
	for (let j = 1; j <= holdout; j++) {
		const t = train + j - 1;
		const e = (y[t] as number) - (level + j * trend + (m >= 2 ? (season[t % m] ?? 0) : 0));
		sse += e * e;
	}
	return sse;
}
