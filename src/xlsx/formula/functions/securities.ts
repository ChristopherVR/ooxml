// Discount securities and Treasury bills: DISC, PRICEDISC, YIELDDISC, INTRATE, RECEIVED,
// PRICEMAT, YIELDMAT and the TBILL functions.
import type { CallContext } from '../context.js';
import { ERR, fail, type Value } from '../values.js';
import { serialArg, yearFrac } from './date-calc.js';
import { int, num, spec } from './helpers.js';
import type { FunctionSpec } from './types.js';

const C = 'Financial';

const basisOf = (args: Value[], index: number): number => {
	const basis = args.length > index && args[index] !== null ? int(args[index]) : 0;
	if (basis < 0 || basis > 4) fail(ERR.NUM);
	return basis;
};

/** Settlement and maturity, which must be in order. */
function span(args: Value[]): { s: number; m: number } {
	const [s, m] = [Math.floor(serialArg(args[0])), Math.floor(serialArg(args[1]))];
	if (s >= m) fail(ERR.NUM);
	return { s, m };
}

const positive = (...values: number[]): void => {
	for (const v of values) if (!(v > 0)) fail(ERR.NUM);
};

/** A security with settlement, maturity, two amounts and an optional basis. */
const discount = (
	name: string,
	description: string,
	syntax: string,
	fn: (a: number, b: number, fraction: number) => number,
): FunctionSpec =>
	spec(name, C, syntax, description, 4, 5, (args, ctx: CallContext) => {
		const { s, m } = span(args);
		const [a, b] = [num(args[2]), num(args[3])];
		positive(a, b);
		return fn(a, b, yearFrac(s, m, basisOf(args, 4), ctx));
	});

/** A security with an issue date: `settlement, maturity, issue, rate, yield-or-price, [basis]`. */
const atMaturity = (
	name: string,
	description: string,
	syntax: string,
	fn: (
		rate: number,
		other: number,
		issueToMaturity: number,
		issueToSettle: number,
		settleToMaturity: number,
	) => number,
): FunctionSpec =>
	spec(name, C, syntax, description, 5, 6, (args, ctx: CallContext) => {
		const { s, m } = span(args);
		const issue = Math.floor(serialArg(args[2]));
		const [rate, other] = [num(args[3]), num(args[4])];
		const basis = basisOf(args, 5);
		if (issue > s || rate < 0 || other < 0) fail(ERR.NUM);
		return fn(
			rate,
			other,
			yearFrac(issue, m, basis, ctx),
			yearFrac(issue, s, basis, ctx),
			yearFrac(s, m, basis, ctx),
		);
	});

/** Days to maturity of a Treasury bill, which may not exceed a year. */
function billDays(args: Value[]): number {
	const { s, m } = span(args);
	const days = m - s;
	if (days > 365) fail(ERR.NUM);
	return days;
}

export const SECURITY_FUNCTIONS: FunctionSpec[] = [
	discount(
		'DISC',
		'The discount rate of a security.',
		'DISC(settlement, maturity, pr, redemption, [basis])',
		(pr, redemption, f) => (redemption - pr) / redemption / f,
	),
	discount(
		'PRICEDISC',
		'The price per 100 face value of a discounted security.',
		'PRICEDISC(settlement, maturity, discount, redemption, [basis])',
		(rate, redemption, f) => redemption - rate * redemption * f,
	),
	discount(
		'YIELDDISC',
		'The annual yield of a discounted security.',
		'YIELDDISC(settlement, maturity, pr, redemption, [basis])',
		(pr, redemption, f) => (redemption - pr) / pr / f,
	),
	discount(
		'INTRATE',
		'The interest rate of a fully invested security.',
		'INTRATE(settlement, maturity, investment, redemption, [basis])',
		(investment, redemption, f) => (redemption - investment) / investment / f,
	),
	discount(
		'RECEIVED',
		'The amount received at maturity for a fully invested security.',
		'RECEIVED(settlement, maturity, investment, discount, [basis])',
		(investment, rate, f) => investment / (1 - rate * f),
	),
	atMaturity(
		'PRICEMAT',
		'The price per 100 face value of a security that pays interest at maturity.',
		'PRICEMAT(settlement, maturity, issue, rate, yld, [basis])',
		(rate, yld, dim, a, dsm) => (100 + dim * rate * 100) / (1 + dsm * yld) - a * rate * 100,
	),
	atMaturity(
		'YIELDMAT',
		'The annual yield of a security that pays interest at maturity.',
		'YIELDMAT(settlement, maturity, issue, rate, pr, [basis])',
		(rate, price, dim, a, dsm) =>
			(1 + dim * rate - (price / 100 + a * rate)) / (price / 100 + a * rate) / dsm,
	),
	spec(
		'TBILLPRICE',
		C,
		'TBILLPRICE(settlement, maturity, discount)',
		'The price per 100 face value of a Treasury bill.',
		3,
		3,
		(args) => {
			const days = billDays(args);
			const rate = num(args[2]);
			positive(rate);
			const price = 100 * (1 - (rate * days) / 360);
			return price > 0 ? price : fail(ERR.NUM);
		},
	),
	spec(
		'TBILLYIELD',
		C,
		'TBILLYIELD(settlement, maturity, pr)',
		'The yield of a Treasury bill.',
		3,
		3,
		(args) => {
			const days = billDays(args);
			const price = num(args[2]);
			positive(price);
			return ((100 - price) / price) * (360 / days);
		},
	),
	spec(
		'TBILLEQ',
		C,
		'TBILLEQ(settlement, maturity, discount)',
		'The bond-equivalent yield of a Treasury bill.',
		3,
		3,
		(args) => {
			const days = billDays(args);
			const rate = num(args[2]);
			positive(rate);
			if (days <= 182) return (365 * rate) / (360 - rate * days);
			const price = 100 - rate * 100 * (days / 360);
			const t = days / 365;
			const root = Math.sqrt(t * t - (2 * t - 1) * (1 - 100 / price));
			return (-t + root) / (t - 0.5);
		},
	),
];
