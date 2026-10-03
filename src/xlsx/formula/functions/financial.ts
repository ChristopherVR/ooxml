import type { CallContext } from '../context.js';
import { ERR, fail, type Value } from '../values.js';
import { collectNumbers, num, numeric, optNum, spec } from './helpers.js';
import type { FunctionSpec } from './types.js';
import { DEPRECIATION_FUNCTIONS } from './depreciation.js';
import { fv, ipmt, pmt, pv, solve } from './financial-core.js';

const C = 'Financial';

const npv = (rate: number, values: number[]): number =>
	values.reduce((acc, v, i) => acc + v / Math.pow(1 + rate, i + 1), 0);

function datedFlows(ctx: CallContext, args: Value[]): { values: number[]; dates: number[] } {
	const values = collectNumbers(ctx, [args[0] ?? null]);
	const dates = collectNumbers(ctx, [args[1] ?? null]).map(Math.floor);
	if (values.length !== dates.length || values.length === 0) fail(ERR.NUM);
	const first = dates[0] as number;
	if (dates.some((d) => d < first)) fail(ERR.NUM);
	return { values, dates };
}

const xnpv = (rate: number, values: number[], dates: number[]): number =>
	values.reduce(
		(acc, v, i) =>
			acc + v / Math.pow(1 + rate, ((dates[i] as number) - (dates[0] as number)) / 365),
		0,
	);

export const FINANCIAL_FUNCTIONS: FunctionSpec[] = [
	...DEPRECIATION_FUNCTIONS,
	numeric(
		'PMT',
		C,
		'PMT(rate, nper, pv, [fv], [type])',
		'The payment for a loan.',
		3,
		5,
		(r, n, p, f, t) => pmt(r, n ?? 0, p ?? 0, f, t),
	),
	numeric(
		'FV',
		C,
		'FV(rate, nper, pmt, [pv], [type])',
		'The future value of an investment.',
		3,
		5,
		(r, n, p, v, t) => fv(r, n ?? 0, p ?? 0, v, t),
	),
	numeric(
		'PV',
		C,
		'PV(rate, nper, pmt, [fv], [type])',
		'The present value of an investment.',
		3,
		5,
		(r, n, p, f, t) => pv(r, n ?? 0, p ?? 0, f, t),
	),
	numeric(
		'NPER',
		C,
		'NPER(rate, pmt, pv, [fv], [type])',
		'The number of periods for an investment.',
		3,
		5,
		(r, p, v, f = 0, t = 0) => {
			const payment = p ?? 0;
			const present = v ?? 0;
			if (r === 0) return payment === 0 ? fail(ERR.NUM) : -(present + f) / payment;
			const num1 = payment * (1 + r * t) - f * r;
			const den = present * r + payment * (1 + r * t);
			if (num1 / den <= 0) fail(ERR.NUM);
			return Math.log(num1 / den) / Math.log(1 + r);
		},
	),
	numeric(
		'IPMT',
		C,
		'IPMT(rate, per, nper, pv, [fv], [type])',
		'The interest part of a payment.',
		4,
		6,
		(r, per, n, p, f, t) => ipmt(r, per ?? 0, n ?? 0, p ?? 0, f, t),
	),
	numeric(
		'PPMT',
		C,
		'PPMT(rate, per, nper, pv, [fv], [type])',
		'The principal part of a payment.',
		4,
		6,
		(r, per, n, p, f, t) => pmt(r, n ?? 0, p ?? 0, f, t) - ipmt(r, per ?? 0, n ?? 0, p ?? 0, f, t),
	),
	numeric(
		'RATE',
		C,
		'RATE(nper, pmt, pv, [fv], [type], [guess])',
		'The interest rate per period.',
		3,
		6,
		(n, p, v, f = 0, t = 0, g = 0.1) =>
			solve((r) => {
				if (Math.abs(r) < 1e-12) return (v ?? 0) + (p ?? 0) * n + f;
				// expm1/log1p keep (1 + r)^n - 1 accurate near r = 0, so a zero rate converges.
				const growth = Math.expm1(n * Math.log1p(r));
				return (v ?? 0) * (1 + growth) + ((p ?? 0) * (1 + r * t) * growth) / r + f;
			}, g),
	),
	spec(
		'NPV',
		C,
		'NPV(rate, value1, [value2], ...)',
		'The net present value of cash flows.',
		2,
		255,
		(args, ctx) => npv(num(args[0]), collectNumbers(ctx, args.slice(1))),
		['value', 'any'],
	),
	spec(
		'IRR',
		C,
		'IRR(values, [guess])',
		'The internal rate of return.',
		1,
		2,
		(args, ctx) => {
			const values = collectNumbers(ctx, [args[0] ?? null]);
			if (!values.some((v) => v > 0) || !values.some((v) => v < 0)) fail(ERR.NUM);
			return solve(
				(r) => values.reduce((acc, v, i) => acc + v / Math.pow(1 + r, i), 0),
				optNum(args, 1, 0.1),
			);
		},
		['any', 'value'],
	),
	spec(
		'XNPV',
		C,
		'XNPV(rate, values, dates)',
		'The net present value of dated cash flows.',
		3,
		3,
		(args, ctx) => {
			const { values, dates } = datedFlows(ctx, args.slice(1));
			return xnpv(num(args[0]), values, dates);
		},
		['value', 'any'],
	),
	spec(
		'XIRR',
		C,
		'XIRR(values, dates, [guess])',
		'The internal rate of return of dated cash flows.',
		2,
		3,
		(args, ctx) => {
			const { values, dates } = datedFlows(ctx, args);
			if (!values.some((v) => v > 0) || !values.some((v) => v < 0)) fail(ERR.NUM);
			return solve((r) => xnpv(r, values, dates), optNum(args, 2, 0.1));
		},
		['any', 'any', 'value'],
	),
	spec(
		'MIRR',
		C,
		'MIRR(values, finance_rate, reinvest_rate)',
		'The modified internal rate of return.',
		3,
		3,
		(args, ctx) => {
			const values = collectNumbers(ctx, [args[0] ?? null]);
			const finance = num(args[1]);
			const reinvest = num(args[2]);
			const n = values.length;
			const positive = values.map((v) => (v > 0 ? v : 0));
			const negative = values.map((v) => (v < 0 ? v : 0));
			const fvPos = positive.reduce((acc, v, i) => acc + v * Math.pow(1 + reinvest, n - 1 - i), 0);
			const pvNeg = negative.reduce((acc, v, i) => acc + v / Math.pow(1 + finance, i), 0);
			if (fvPos === 0 || pvNeg === 0) fail(ERR.DIV0);
			return Math.pow(-fvPos / pvNeg, 1 / (n - 1)) - 1;
		},
		['any', 'value', 'value'],
	),
];
