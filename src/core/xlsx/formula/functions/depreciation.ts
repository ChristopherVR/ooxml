import { ERR, fail } from '../values';
import { bool, collectNumbers, num, numeric, spec } from './helpers';
import type { FunctionSpec } from './types';

import { cumulative } from './financial-core';

const C = 'Financial';

/** Depreciation, rate conversion and price-format functions. */
/**
 * Declining-balance depreciation over `[start, end]`, taking each whole period in turn and weighting
 * it by how much of it the span covers. Unless `noSwitch`, a period uses straight-line depreciation
 * of the remaining book value once that exceeds the declining-balance amount.
 */
function variableDeclining(
	cost: number,
	salvage: number,
	life: number,
	start: number,
	end: number,
	factor: number,
	noSwitch: boolean,
): number {
	let book = cost;
	let total = 0;
	for (let period = 1; period - 1 < end; period++) {
		const declining = (book * factor) / life;
		const straight = (book - salvage) / Math.max(life - (period - 1), 1);
		const wanted = noSwitch ? declining : Math.max(declining, straight);
		const dep = Math.max(0, Math.min(wanted, book - salvage));
		const covered = Math.min(end, period) - Math.max(start, period - 1);
		if (covered > 0) total += dep * Math.min(covered, 1);
		book -= dep;
	}
	return total;
}

export const DEPRECIATION_FUNCTIONS: FunctionSpec[] = [
	numeric('SLN', C, 'SLN(cost, salvage, life)', 'Straight-line depreciation.', 3, 3, (c, s, l) =>
		l === 0 ? fail(ERR.DIV0) : (c - (s ?? 0)) / (l ?? 1),
	),
	numeric(
		'SYD',
		C,
		'SYD(cost, salvage, life, per)',
		"Sum-of-years' digits depreciation.",
		4,
		4,
		(c, s = 0, l = 1, p = 1) => {
			if (l <= 0 || p <= 0 || p > l) fail(ERR.NUM);
			return ((c - s) * (l - p + 1) * 2) / (l * (l + 1));
		},
	),
	numeric(
		'DDB',
		C,
		'DDB(cost, salvage, life, period, [factor])',
		'Double-declining balance depreciation.',
		4,
		5,
		(c, s = 0, l = 1, p = 1, f = 2) => {
			if (c < 0 || s < 0 || l <= 0 || p <= 0 || p > l || f <= 0) fail(ERR.NUM);
			// Native Excel retains fractional periods above one and applies the
			// first-period amount below one. The closed form also bounds work for
			// large life/period arguments, instead of iterating over every period.
			const rate = Math.min(f / l, 1);
			const book = c * Math.pow(1 - rate, Math.max(0, p - 1));
			return Math.max(0, Math.min(book * rate, book - s));
		},
	),
	spec(
		'VDB',
		C,
		'VDB(cost, salvage, life, start_period, end_period, [factor], [no_switch])',
		'Depreciation over any span of periods, switching to straight-line when that is larger.',
		5,
		7,
		(args) => {
			const [cost, salvage, life, start, end] = [0, 1, 2, 3, 4].map((i) => num(args[i]));
			const factor = args.length > 5 && args[5] !== null ? num(args[5]) : 2;
			const noSwitch = args.length > 6 && args[6] !== null ? bool(args[6]) : false;
			if (
				[cost, salvage, life, start, end, factor].some((v) => v === undefined || Number.isNaN(v)) ||
				(cost as number) < 0 ||
				(salvage as number) < 0 ||
				(life as number) <= 0 ||
				(start as number) < 0 ||
				(end as number) < (start as number) ||
				(end as number) > (life as number) ||
				factor <= 0
			)
				fail(ERR.NUM);
			return variableDeclining(
				cost as number,
				salvage as number,
				life as number,
				start as number,
				end as number,
				factor,
				noSwitch,
			);
		},
	),
	numeric(
		'DB',
		C,
		'DB(cost, salvage, life, period, [month])',
		'Fixed-declining balance depreciation.',
		4,
		5,
		(c, s = 0, l = 1, p = 1, m = 12) => {
			// A partial first year (month < 12) adds a final period after the life.
			const month = Math.trunc(m),
				period = Math.trunc(p);
			if (
				c < 0 ||
				s < 0 ||
				l <= 0 ||
				p <= 0 ||
				month < 1 ||
				month > 12 ||
				period > (month < 12 ? l + 1 : l)
			)
				fail(ERR.NUM);
			if (c === 0) return 0;
			const rate = Math.round((1 - Math.pow(s / c, 1 / l)) * 1000) / 1000;
			const first = (c * rate * month) / 12;
			if (period <= 1) return first;
			const book = (c - first) * Math.pow(1 - rate, period - 2);
			return period === Math.trunc(l) + 1 ? (book * rate * (12 - month)) / 12 : book * rate;
		},
	),
	numeric(
		'EFFECT',
		C,
		'EFFECT(nominal_rate, npery)',
		'The effective annual interest rate.',
		2,
		2,
		(r, n) => {
			const p = Math.trunc(n ?? 0);
			if (r <= 0 || p < 1) fail(ERR.NUM);
			return Math.pow(1 + r / p, p) - 1;
		},
	),
	numeric(
		'NOMINAL',
		C,
		'NOMINAL(effect_rate, npery)',
		'The nominal annual interest rate.',
		2,
		2,
		(r, n) => {
			const p = Math.trunc(n ?? 0);
			if (r <= 0 || p < 1) fail(ERR.NUM);
			return p * (Math.pow(1 + r, 1 / p) - 1);
		},
	),
	numeric(
		'CUMIPMT',
		C,
		'CUMIPMT(rate, nper, pv, start_period, end_period, type)',
		'Cumulative interest paid between periods.',
		6,
		6,
		(...a) => cumulative(a, false),
	),
	numeric(
		'CUMPRINC',
		C,
		'CUMPRINC(rate, nper, pv, start_period, end_period, type)',
		'Cumulative principal paid between periods.',
		6,
		6,
		(...a) => cumulative(a, true),
	),
	numeric(
		'ISPMT',
		C,
		'ISPMT(rate, per, nper, pv)',
		'Interest paid in a period of a straight-line loan.',
		4,
		4,
		(r, per = 0, n = 1, p = 0) => {
			if (n === 0) fail(ERR.DIV0);
			return p * r * (per / n - 1);
		},
	),
	numeric(
		'PDURATION',
		C,
		'PDURATION(rate, pv, fv)',
		'Periods required to reach a value.',
		3,
		3,
		(r, p = 0, f = 0) => {
			if (r <= 0 || p <= 0 || f <= 0) fail(ERR.NUM);
			return (Math.log(f) - Math.log(p)) / Math.log(1 + r);
		},
	),
	numeric(
		'RRI',
		C,
		'RRI(nper, pv, fv)',
		'An equivalent interest rate for growth.',
		3,
		3,
		(n, p = 0, f = 0) => {
			if (n <= 0 || p === 0) fail(ERR.NUM);
			return Math.pow(f / p, 1 / n) - 1;
		},
	),
	spec(
		'FVSCHEDULE',
		C,
		'FVSCHEDULE(principal, schedule)',
		'Future value with a schedule of rates.',
		2,
		2,
		(args, ctx) =>
			collectNumbers(ctx, [args[1] ?? null]).reduce((acc, r) => acc * (1 + r), num(args[0])),
		['value', 'any'],
	),
	numeric(
		'DOLLARDE',
		C,
		'DOLLARDE(fractional_dollar, fraction)',
		'Converts a fractional price to a decimal.',
		2,
		2,
		(d, f) => {
			const frac = Math.trunc(f ?? 0);
			if (frac < 0) fail(ERR.NUM);
			if (frac === 0) fail(ERR.DIV0);
			const digits = Math.ceil(Math.log10(frac));
			const int = Math.trunc(d);
			return int + ((d - int) * Math.pow(10, digits)) / frac;
		},
	),
	numeric(
		'DOLLARFR',
		C,
		'DOLLARFR(decimal_dollar, fraction)',
		'Converts a decimal price to a fraction.',
		2,
		2,
		(d, f) => {
			const frac = Math.trunc(f ?? 0);
			if (frac < 0) fail(ERR.NUM);
			if (frac === 0) fail(ERR.DIV0);
			const digits = Math.ceil(Math.log10(frac));
			const int = Math.trunc(d);
			return int + ((d - int) * frac) / Math.pow(10, digits);
		},
	),
];
