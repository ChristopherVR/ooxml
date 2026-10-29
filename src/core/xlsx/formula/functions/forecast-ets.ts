// FORECAST.ETS, FORECAST.ETS.CONFINT, FORECAST.ETS.SEASONALITY and FORECAST.ETS.STAT. Built on
// the additive exponential triple smoothing in `ets-core`, whose filter, starting values, error
// statistics and interval were identified against Excel's recorded answers
// (forecast-ets-excel.test.ts). Excel's own search for the smoothing parameters is undocumented
// and not reproduced exactly (see the deviations in functions.md).
import type { CallContext } from '../context';
import { ERR, fail, type Value } from '../values';
import { normInv } from './dist-core';
import { etsErrorSd, type EtsFit, etsHoldout, etsPredict, fitEts } from './ets-core';
import { detectSeasonality } from './ets-season';
import { etsSeries } from './ets-input';
import { num, omitted, spec } from './helpers';
import type { FunctionSpec } from './types';

const C = 'Statistical';
const MAX_SEASON = 8760;

/** An optional numeric argument; an omitted or empty slot takes the default 1. */
const opt = (args: Value[], i: number): number => (omitted(args, i) ? 1 : num(args[i]));

interface Model {
	fit: EtsFit;
	position: (time: number) => number;
	step: number;
	y: number[];
}

/** Builds the model for a call; `at` is the index of the optional seasonality argument. */
function model(ctx: CallContext, args: Value[], at: number): Model {
	const seasonality = opt(args, at);
	const completion = opt(args, at + 1);
	const aggregation = opt(args, at + 2);
	if (!Number.isInteger(seasonality) || seasonality < 0 || seasonality > MAX_SEASON) fail(ERR.NUM);
	const { y, position, step } = etsSeries(
		ctx,
		args[1] ?? null,
		args[2] ?? null,
		completion,
		aggregation,
	);
	// 1 means "detect"; 0 turns the seasonal component off.
	const m = seasonality === 1 ? detectSeasonality(y) : seasonality;
	return { fit: fitEts(y, m), position, step, y };
}

/** Steps from the last point to the target date, rounded to the grid. */
function stepsAhead(m: Model, target: number): number {
	const k = Math.round(m.position(target));
	if (k < 0) fail(ERR.NUM);
	return k - (m.fit.n - 1);
}

const MAE = (e: number[]): number => e.reduce((a, v) => a + Math.abs(v), 0) / e.length;

function statistic(m: Model, kind: number): number {
	const { fit, y } = m;
	// Excel measures the one-step errors of the last `holdout` points only (every point of a short
	// series); the scale of MASE comes from the points before them.
	const holdout = etsHoldout(y.length, fit.m);
	const from = y.length - holdout;
	const errors = y.map((v, i) => v - (fit.fitted[i] as number)).slice(holdout ? from : 0);
	switch (kind) {
		case 1:
			return fit.alpha;
		case 2:
			return fit.beta;
		case 3:
			return fit.gamma;
		case 4: {
			const naive = y.slice(1, from).map((v, i) => v - (y[i] as number));
			const scale = naive.length ? MAE(naive) : 0;
			// A series that never changes has no scale; Excel reports 0 when it is also fitted exactly.
			if (scale === 0) return MAE(errors) === 0 ? 0 : fail(ERR.DIV0);
			return MAE(errors) / scale;
		}
		case 5:
			return (
				errors.reduce((a, e, i) => {
					const at = (holdout ? from : 0) + i;
					const d = Math.abs(y[at] as number) + Math.abs(fit.fitted[at] as number);
					return a + (d === 0 ? 0 : (2 * Math.abs(e)) / d);
				}, 0) / errors.length
			);
		case 6:
			return MAE(errors);
		case 7:
			return Math.sqrt(errors.reduce((a, e) => a + e * e, 0) / errors.length);
		case 8:
			return m.step;
		default:
			return fail(ERR.NUM);
	}
}

export const FORECAST_ETS_FUNCTIONS: FunctionSpec[] = [
	spec(
		'FORECAST.ETS',
		C,
		'FORECAST.ETS(target_date, values, timeline, [seasonality], [data_completion], [aggregation])',
		'Predicts a future value with additive exponential triple smoothing (AAA).',
		3,
		6,
		(args, ctx) => {
			const target = num(args[0]);
			const built = model(ctx, args, 3);
			return etsPredict(built.fit, stepsAhead(built, target));
		},
		['value', 'any', 'any', 'value'],
	),
	spec(
		'FORECAST.ETS.CONFINT',
		C,
		'FORECAST.ETS.CONFINT(target_date, values, timeline, [confidence_level], [seasonality], [data_completion], [aggregation])',
		'The confidence interval half-width of a FORECAST.ETS prediction.',
		3,
		7,
		(args, ctx) => {
			const target = num(args[0]);
			const level = omitted(args, 3) ? 0.95 : num(args[3]);
			if (!(level > 0 && level < 1)) fail(ERR.NUM);
			const rest = [...args.slice(0, 3), ...args.slice(4)];
			const built = model(ctx, rest, 3);
			const h = stepsAhead(built, target);
			return normInv((1 + level) / 2) * etsErrorSd(built.fit, h);
		},
		['value', 'any', 'any', 'value'],
	),
	spec(
		'FORECAST.ETS.SEASONALITY',
		C,
		'FORECAST.ETS.SEASONALITY(values, timeline, [data_completion], [aggregation])',
		'The length of the repeating pattern detected in a series (0 when there is none).',
		2,
		4,
		(args, ctx) => {
			const completion = opt(args, 2);
			const aggregation = opt(args, 3);
			const { y } = etsSeries(ctx, args[0] ?? null, args[1] ?? null, completion, aggregation);
			return detectSeasonality(y);
		},
		['any', 'any', 'value'],
	),
	spec(
		'FORECAST.ETS.STAT',
		C,
		'FORECAST.ETS.STAT(values, timeline, statistic_type, [seasonality], [data_completion], [aggregation])',
		'A statistic of the exponential smoothing model: alpha, beta, gamma, MASE, SMAPE, MAE, RMSE or step size.',
		3,
		6,
		(args, ctx) => {
			const kind = num(args[2]);
			if (!Number.isInteger(kind) || kind < 1 || kind > 8) fail(ERR.NUM);
			// The shared model reads values and timeline from positions 1 and 2.
			const shifted = [null, args[0] ?? null, args[1] ?? null, ...args.slice(3)];
			return statistic(model(ctx, shifted, 3), kind);
		},
		['any', 'any', 'value'],
	),
];
