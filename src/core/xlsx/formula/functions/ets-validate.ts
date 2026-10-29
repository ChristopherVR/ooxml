// The error Excel's FORECAST.ETS search minimises (identified from 1100 probe series, see
// ets-search.ts and docs/xlsx-parity-status.md): the series is split into a training part and the
// hold-out of `etsHoldout` trailing points. The filter runs over the training part one step at a
// time, then forecasts the hold-out several steps ahead without updating. The cost is the sum of
// the squared one-step errors of the training part plus the squared multi-step forecast errors of
// the hold-out. Series too short for a hold-out use the plain one-step error.
import type { EtsInit } from './ets-init';

interface Params {
	alpha: number;
	beta: number;
	gamma: number;
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
