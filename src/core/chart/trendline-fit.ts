// Fitting a chart trendline in data space: the regression of one kind over a series' x and y
// values, sampled along the data range plus the forward and backward extensions, with Office's
// equation text and R-squared. A painter maps the samples to pixels. Category charts fit against
// the category positions 1..n (as Excel does); scatter charts against the series' own x values.
import type { ChartTrendlineKind } from './model-analysis';
import { computeLinearRegression, computeRSquared, fitPolynomial } from './regression';
import {
	formatExponentialEquation,
	formatLinearEquation,
	formatLogarithmicEquation,
	formatPolynomialEquation,
	formatPowerEquation,
} from './trendline-equation';

/** What a fit needs from a trendline: its kind and options (`ChartTrendline`, resolved). */
export interface ChartTrendlineFitOptions {
	kind: ChartTrendlineKind;
	/** Polynomial order, capped at 6 (default 2). */
	order?: number | undefined;
	/** Moving-average period (default 2). */
	period?: number | undefined;
	forward?: number | undefined;
	backward?: number | undefined;
	/** Linear: the y value the line is forced through at x = 0. */
	intercept?: number | undefined;
}

/** A point of a fitted trendline, in data space. */
export interface ChartTrendlinePoint {
	x: number;
	y: number;
}

/** A fitted trendline. */
export interface ChartTrendlineFit {
	/** Samples in data space; empty when the data is too sparse or the kind cannot fit it. */
	points: ChartTrendlinePoint[];
	/** Office's equation text (`y = 2x + 1`), or the moving-average description. */
	equation: string;
	rSquared: number;
	/**
	 * The fitted coefficients: linear and logarithmic `[intercept, slope]`, exponential and power
	 * `[a, b]`, polynomial ascending by power; empty for a moving average or no fit.
	 */
	coefficients: number[];
}

const noFit = (): ChartTrendlineFit => ({
	points: [],
	equation: '',
	rSquared: 0,
	coefficients: [],
});

interface Model {
	evaluate: (x: number) => number;
	equation: string;
	rSquared: number;
	coefficients: number[];
}

function fitModel(
	options: ChartTrendlineFitOptions,
	xVals: readonly number[],
	yVals: readonly number[],
): Model | undefined {
	const xs = [...xVals];
	const ys = [...yVals];
	switch (options.kind) {
		case 'linear': {
			const reg = computeLinearRegression(xs, ys);
			const fixed = options.intercept;
			const slope =
				fixed !== undefined
					? ys.reduce((s, y, i) => s + (y - fixed) * xs[i]!, 0) / xs.reduce((s, x) => s + x * x, 0)
					: reg.slope;
			const b = fixed ?? reg.intercept;
			return {
				evaluate: (x) => slope * x + b,
				equation: formatLinearEquation(slope, b),
				rSquared: reg.rSquared,
				coefficients: [b, slope],
			};
		}
		case 'exponential': {
			const logY = ys.filter((y) => y > 0).map((y) => Math.log(y));
			const posX = xs.filter((_x, i) => ys[i]! > 0);
			if (logY.length < 2) return undefined;
			const reg = computeLinearRegression(posX, logY);
			const a = Math.exp(reg.intercept);
			const b = reg.slope;
			return {
				evaluate: (x) => a * Math.exp(b * x),
				equation: formatExponentialEquation(a, b),
				rSquared: reg.rSquared,
				coefficients: [a, b],
			};
		}
		case 'logarithmic': {
			const lnX = xs.filter((x) => x > 0).map((x) => Math.log(x));
			const posY = ys.filter((_y, i) => xs[i]! > 0);
			if (lnX.length < 2) return undefined;
			const reg = computeLinearRegression(lnX, posY);
			return {
				evaluate: (x) => (x > 0 ? reg.slope * Math.log(x) + reg.intercept : 0),
				equation: formatLogarithmicEquation(reg.slope, reg.intercept),
				rSquared: reg.rSquared,
				coefficients: [reg.intercept, reg.slope],
			};
		}
		case 'power': {
			const logX = xs.filter((x, i) => x > 0 && ys[i]! > 0).map((x) => Math.log(x));
			const logY = ys.filter((y, i) => y > 0 && xs[i]! > 0).map((y) => Math.log(y));
			if (logX.length < 2) return undefined;
			const reg = computeLinearRegression(logX, logY);
			const a = Math.exp(reg.intercept);
			return {
				evaluate: (x) => (x > 0 ? a * x ** reg.slope : 0),
				equation: formatPowerEquation(a, reg.slope),
				rSquared: reg.rSquared,
				coefficients: [a, reg.slope],
			};
		}
		case 'polynomial': {
			const order = Math.min(options.order ?? 2, 6);
			const coeffs = fitPolynomial(xs, ys, order);
			const evaluate = (x: number) => coeffs.reduce((s, c, i) => s + c * x ** i, 0);
			return {
				evaluate,
				equation: formatPolynomialEquation(coeffs),
				rSquared: computeRSquared(xs, ys, evaluate),
				coefficients: coeffs,
			};
		}
		default:
			return undefined;
	}
}

/**
 * Fits one trendline over paired values (`xVals[i]`, `yVals[i]`), both finite and of equal length.
 * A moving average yields one point per full period at the data's own x values; every other kind
 * is sampled at least 20 times (4 per x unit) from the smallest x less `backward` to the largest
 * plus `forward`, skipping samples where the curve is not finite.
 */
export function fitChartTrendline(
	options: ChartTrendlineFitOptions,
	xVals: readonly number[],
	yVals: readonly number[],
): ChartTrendlineFit {
	const count = yVals.length;
	if (count < 2) return noFit();
	if (options.kind === 'movingAvg') {
		const period = options.period ?? 2;
		const points: ChartTrendlinePoint[] = [];
		for (let i = period - 1; i < count; i++) {
			let sum = 0;
			for (let j = i - period + 1; j <= i; j++) sum += yVals[j]!;
			points.push({ x: xVals[i]!, y: sum / period });
		}
		return { points, equation: `${period}-period moving average`, rSquared: 0, coefficients: [] };
	}
	const model = fitModel(options, xVals, yVals);
	if (!model) return noFit();
	const startX = Math.min(...xVals) - (options.backward ?? 0);
	const endX = Math.max(...xVals) + (options.forward ?? 0);
	const steps = Math.max(Math.ceil((endX - startX) * 4), 20);
	const points: ChartTrendlinePoint[] = [];
	for (let step = 0; step <= steps; step++) {
		const x = startX + ((endX - startX) * step) / steps;
		const y = model.evaluate(x);
		if (Number.isFinite(y)) points.push({ x, y });
	}
	return {
		points,
		equation: model.equation,
		rSquared: model.rSquared,
		coefficients: model.coefficients,
	};
}
