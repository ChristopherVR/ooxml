// ODDFPRICE, ODDFYIELD, ODDLPRICE and ODDLYIELD: securities with an odd first or last period.
// Odd periods are split into quasi-coupon periods (the standard long/short generalisation).
import type { CallContext } from '../context';
import { addMonths, days360, serialArg } from './date-calc';
import { daysInMonth, serialToYmd } from '../date-serial';
import { ERR, fail, type Value } from '../values';
import { quasiPeriods } from './bonds';
import { solve } from './financial-core';
import { int, num, spec } from './helpers';
import type { FunctionSpec } from './types';

const C = 'Financial';

interface Calendar {
	/** Fraction of coupon periods (quasi periods) between two dates within `[from, to]`. */
	fraction(a: number, b: number): number;
}

/** Day counting of one basis over the quasi periods `dates`. */
function calendar(dates: number[], frequency: number, basis: number, ctx: CallContext): Calendar {
	const count = (a: number, b: number) =>
		basis === 0 || basis === 4 ? days360(a, b, basis === 4, ctx) : b - a;
	return {
		fraction(a, b) {
			let sum = 0;
			for (let i = 0; i + 1 < dates.length; i++) {
				const [p, q] = [dates[i] as number, dates[i + 1] as number];
				const days = count(Math.max(a, p), Math.min(b, q));
				if (days <= 0) continue;
				sum += days / (basis === 1 ? q - p : basis === 3 ? 365 / frequency : 360 / frequency);
			}
			return sum;
		},
	};
}

function basisOf(args: Value[], index: number): number {
	const basis = args.length > index && args[index] !== null ? int(args[index]) : 0;
	if (basis < 0 || basis > 4) fail(ERR.NUM);
	return basis;
}

function frequencyOf(value: Value | undefined): number {
	const frequency = int(value);
	if (![1, 2, 4].includes(frequency)) fail(ERR.NUM);
	return frequency;
}

interface First {
	settlement: number;
	maturity: number;
	issue: number;
	firstCoupon: number;
	rate: number;
	redemption: number;
	frequency: number;
	basis: number;
}

/** Price per 100 of a security with an odd first period at yield `yld`. */
function oddFirstPrice(s: First, yld: number, ctx: CallContext): number {
	const { settlement, maturity, issue, firstCoupon, rate, redemption, frequency, basis } = s;
	const months = 12 / frequency;
	const coupon = (100 * rate) / frequency;
	const quasi = calendar(
		quasiPeriods(firstCoupon, issue, firstCoupon, frequency, ctx),
		frequency,
		basis,
		ctx,
	);
	// Time from settlement to the first coupon, in coupon periods.
	const toFirst = quasi.fraction(settlement, firstCoupon);
	const { year, month, day } = serialToYmd(firstCoupon, ctx.date1904);
	const monthEnd = day === daysInMonth(year, month);
	let n = 1;
	while (addMonths(firstCoupon, n * months, ctx, monthEnd) < maturity) n += 1;
	n += 1;
	const base = 1 + yld / frequency;
	let price = redemption / base ** (toFirst + n - 1);
	price += (coupon * quasi.fraction(issue, firstCoupon)) / base ** toFirst;
	for (let k = 2; k <= n; k++) price += coupon / base ** (toFirst + k - 1);
	return price - coupon * quasi.fraction(issue, settlement);
}

function firstOf(args: Value[]): First {
	const [settlement, maturity, issue, firstCoupon] = [0, 1, 2, 3].map((i) =>
		Math.floor(serialArg(args[i])),
	) as [number, number, number, number];
	const frequency = frequencyOf(args[7]);
	const first: First = {
		settlement,
		maturity,
		issue,
		firstCoupon,
		rate: num(args[4]),
		redemption: num(args[6]),
		frequency,
		basis: basisOf(args, 8),
	};
	if (!(maturity > firstCoupon && firstCoupon > settlement && settlement > issue)) fail(ERR.NUM);
	if (first.rate < 0 || first.redemption <= 0) fail(ERR.NUM);
	return first;
}

interface Last {
	settlement: number;
	maturity: number;
	rate: number;
	redemption: number;
	frequency: number;
	cal: Calendar;
	lastInterest: number;
}

function lastOf(args: Value[], ctx: CallContext): Last {
	const [settlement, maturity, lastInterest] = [0, 1, 2].map((i) =>
		Math.floor(serialArg(args[i])),
	) as [number, number, number];
	const frequency = frequencyOf(args[6]);
	const last: Last = {
		settlement,
		maturity,
		lastInterest,
		rate: num(args[3]),
		redemption: num(args[5]),
		frequency,
		cal: calendar(
			quasiPeriods(lastInterest, lastInterest, maturity, frequency, ctx),
			frequency,
			basisOf(args, 7),
			ctx,
		),
	};
	if (!(maturity > settlement && settlement > lastInterest)) fail(ERR.NUM);
	if (last.rate < 0 || last.redemption <= 0) fail(ERR.NUM);
	return last;
}

/** Price per 100 of a security with an odd last period at yield `yld`. */
function oddLastPrice(l: Last, yld: number): number {
	const coupon = (100 * l.rate) / l.frequency;
	const full = l.redemption + coupon * l.cal.fraction(l.lastInterest, l.maturity);
	const discount = 1 + (yld / l.frequency) * l.cal.fraction(l.settlement, l.maturity);
	return full / discount - coupon * l.cal.fraction(l.lastInterest, l.settlement);
}

export const ODD_SECURITY_FUNCTIONS: FunctionSpec[] = [
	spec(
		'ODDFPRICE',
		C,
		'ODDFPRICE(settlement, maturity, issue, first_coupon, rate, yld, redemption, frequency, [basis])',
		'The price per 100 face value of a security with an odd first period.',
		8,
		9,
		(args, ctx) => {
			const first = firstOf(args);
			const yld = num(args[5]);
			if (yld < 0) fail(ERR.NUM);
			return oddFirstPrice(first, yld, ctx);
		},
	),
	spec(
		'ODDFYIELD',
		C,
		'ODDFYIELD(settlement, maturity, issue, first_coupon, rate, pr, redemption, frequency, [basis])',
		'The yield of a security with an odd first period.',
		8,
		9,
		(args, ctx) => {
			const first = firstOf(args);
			const price = num(args[5]);
			if (price <= 0) fail(ERR.NUM);
			return solve((y) => oddFirstPrice(first, y, ctx) - price, first.rate || 0.05);
		},
	),
	spec(
		'ODDLPRICE',
		C,
		'ODDLPRICE(settlement, maturity, last_interest, rate, yld, redemption, frequency, [basis])',
		'The price per 100 face value of a security with an odd last period.',
		7,
		8,
		(args, ctx) => {
			const last = lastOf(args, ctx);
			const yld = num(args[4]);
			if (yld < 0) fail(ERR.NUM);
			return oddLastPrice(last, yld);
		},
	),
	spec(
		'ODDLYIELD',
		C,
		'ODDLYIELD(settlement, maturity, last_interest, rate, pr, redemption, frequency, [basis])',
		'The yield of a security with an odd last period.',
		7,
		8,
		(args, ctx) => {
			const last = lastOf(args, ctx);
			const price = num(args[4]);
			if (price <= 0) fail(ERR.NUM);
			const coupon = (100 * last.rate) / last.frequency;
			const cal = last.cal;
			const dirty = price + coupon * cal.fraction(last.lastInterest, last.settlement);
			const full = last.redemption + coupon * cal.fraction(last.lastInterest, last.maturity);
			return (
				((full - dirty) / dirty) * (last.frequency / cal.fraction(last.settlement, last.maturity))
			);
		},
	),
];
