import { ERR, fail } from '../values';
import { eliminate } from './matrix';

/** The least-squares fit of `y` on the columns of `x` (observations are rows). */
export interface Fit {
	/** One slope per x column, in column order. */
	slopes: number[];
	/** The intercept (0 when the fit was forced through the origin). */
	intercept: number;
	/** `(X'X)^-1` of the centred columns (or of the raw columns without an intercept). */
	inverse: number[][];
	/** Column means, used for the intercept's standard error. */
	means: number[];
	n: number;
	hasIntercept: boolean;
	ssTotal: number;
	ssResid: number;
}

const total = (values: readonly number[]): number => values.reduce((a, b) => a + b, 0);

/**
 * Ordinary least squares by the normal equations on mean-centred data (which keeps polynomial
 * and offset data well conditioned). Collinear columns are `#NUM!`; Excel instead zeroes the
 * redundant coefficient, which this does not reproduce.
 */
export function fitLeastSquares(x: number[][], y: number[], hasIntercept: boolean): Fit {
	const n = y.length;
	const k = x[0]?.length ?? 0;
	if (n === 0 || k === 0) fail(ERR.VALUE);
	const means = hasIntercept
		? Array.from({ length: k }, (_, j) => total(x.map((row) => row[j] ?? 0)) / n)
		: new Array<number>(k).fill(0);
	const yMean = hasIntercept ? total(y) / n : 0;
	const xc = x.map((row) => row.map((v, j) => v - (means[j] ?? 0)));
	const yc = y.map((v) => v - yMean);
	const gram = Array.from({ length: k }, (_, a) =>
		Array.from({ length: k }, (_, b) => total(xc.map((row) => (row[a] ?? 0) * (row[b] ?? 0)))),
	);
	const moment = Array.from({ length: k }, (_, a) =>
		total(xc.map((row, i) => (row[a] ?? 0) * (yc[i] ?? 0))),
	);
	const { det, inverse } = eliminate(gram, true);
	if (!inverse || det === 0 || !Number.isFinite(det)) fail(ERR.NUM);
	const slopes = inverse.map((row) => total(row.map((v, j) => v * (moment[j] ?? 0))));
	const intercept = hasIntercept ? yMean - total(slopes.map((m, j) => m * (means[j] ?? 0))) : 0;
	let ssResid = 0;
	let ssTotal = 0;
	for (let i = 0; i < n; i++) {
		const predicted = intercept + total(slopes.map((m, j) => m * (x[i]?.[j] ?? 0)));
		ssResid += ((y[i] ?? 0) - predicted) ** 2;
		ssTotal += ((y[i] ?? 0) - yMean) ** 2;
	}
	return { slopes, intercept, inverse, means, n, hasIntercept, ssTotal, ssResid };
}

/** The regression statistics LINEST reports, from a fit. */
export interface FitStats {
	slopeErrors: number[];
	interceptError: number | undefined;
	rSquared: number;
	standardError: number;
	fStatistic: number;
	df: number;
	ssRegression: number;
	ssResidual: number;
}

export function fitStatistics(fit: Fit): FitStats {
	const k = fit.slopes.length;
	const df = fit.n - k - (fit.hasIntercept ? 1 : 0);
	if (df < 1) fail(ERR.NUM);
	const ssRegression = fit.ssTotal - fit.ssResid;
	const variance = fit.ssResid / df;
	const standardError = Math.sqrt(variance);
	const slopeErrors = fit.inverse.map((row, j) => Math.sqrt(variance * (row[j] ?? 0)));
	let interceptError: number | undefined;
	if (fit.hasIntercept) {
		let quad = 0;
		for (let a = 0; a < k; a++)
			for (let b = 0; b < k; b++)
				quad += (fit.means[a] ?? 0) * (fit.inverse[a]?.[b] ?? 0) * (fit.means[b] ?? 0);
		interceptError = Math.sqrt(variance * (1 / fit.n + quad));
	}
	return {
		slopeErrors,
		interceptError,
		rSquared: fit.ssTotal === 0 ? 1 : ssRegression / fit.ssTotal,
		standardError,
		fStatistic: fit.ssResid === 0 ? Infinity : ssRegression / k / variance,
		df,
		ssRegression,
		ssResidual: fit.ssResid,
	};
}
