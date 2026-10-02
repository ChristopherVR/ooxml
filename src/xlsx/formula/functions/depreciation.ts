import { ERR, fail } from '../values.js';
import { collectNumbers, num, numeric, spec } from './helpers.js';
import type { FunctionSpec } from './types.js';

import { cumulative } from './financial-core.js';

const C = 'Financial';

/** Depreciation, rate conversion and price-format functions. */
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
			let total = 0;
			let dep = 0;
			for (let i = 1; i <= Math.ceil(p); i++) {
				dep = Math.max(0, Math.min(((c - total) * f) / l, c - s - total));
				total += dep;
			}
			return dep;
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
			if (c < 0 || s < 0 || l <= 0 || p <= 0 || m < 1 || m > 12 || p > l + 1) fail(ERR.NUM);
			if (c === 0) return 0;
			const rate = Math.round((1 - Math.pow(s / c, 1 / l)) * 1000) / 1000;
			let total = 0;
			let dep = 0;
			for (let i = 1; i <= Math.trunc(p); i++) {
				if (i === 1) dep = (c * rate * m) / 12;
				else if (i === Math.trunc(l) + 1) dep = ((c - total) * rate * (12 - m)) / 12;
				else dep = (c - total) * rate;
				total += dep;
			}
			return dep;
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
