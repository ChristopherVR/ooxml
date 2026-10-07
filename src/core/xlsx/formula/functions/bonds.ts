// Bond pricing and yield (PRICE, YIELD, DURATION, MDURATION) and accrued interest.
import type { CallContext } from '../context';
import { ERR, fail, type Value } from '../values';
import { couponFacts, type CouponFacts } from './coupons';
import { addMonths, days360, serialArg, yearFrac } from './date-calc';
import { solve } from './financial-core';
import { daysInMonth, serialToYmd } from '../date-serial';
import { bool, int, num, optNum, spec } from './helpers';
import type { FunctionSpec } from './types';

const C = 'Financial';

const basisOf = (args: Value[], index: number): number => {
	const basis = args.length > index && args[index] !== null ? int(args[index]) : 0;
	if (basis < 0 || basis > 4) fail(ERR.NUM);
	return basis;
};

/** Present value, per 100 of redemption, of the remaining coupons and the redemption. */
function pricePerHundred(
	facts: CouponFacts,
	rate: number,
	yld: number,
	redemption: number,
	frequency: number,
): number {
	const { count: n, daysSincePrevious: a, daysInPeriod: e, daysToNext: dsc } = facts;
	const coupon = (100 * rate) / frequency;
	const accrued = coupon * (a / e);
	if (n === 1) return (redemption + coupon) / (1 + ((dsc / e) * yld) / frequency) - accrued;
	let price = redemption / (1 + yld / frequency) ** (n - 1 + dsc / e);
	for (let k = 1; k <= n; k++) price += coupon / (1 + yld / frequency) ** (k - 1 + dsc / e);
	return price - accrued;
}

interface Bond {
	facts: CouponFacts;
	rate: number;
	redemption: number;
	frequency: number;
}

function bondOf(args: Value[], ctx: CallContext, redemptionAt: number, basisAt: number): Bond {
	const frequency = int(args[5]);
	const facts = couponFacts(
		serialArg(args[0]),
		serialArg(args[1]),
		frequency,
		basisOf(args, basisAt),
		ctx,
	);
	return { facts, rate: num(args[2]), redemption: num(args[redemptionAt]), frequency };
}

function duration(bond: Bond, yld: number): number {
	const { facts, rate, frequency } = bond;
	const { count: n, daysInPeriod: e, daysToNext: dsc } = facts;
	const coupon = (100 * rate) / frequency;
	let weighted = 0;
	let total = 0;
	for (let k = 1; k <= n; k++) {
		const periods = k - 1 + dsc / e;
		const flow = k === n ? coupon + 100 : coupon;
		const value = flow / (1 + yld / frequency) ** periods;
		weighted += (periods / frequency) * value;
		total += value;
	}
	return weighted / total;
}

/** Quasi-coupon period start dates around `anchor`, covering `from` through `to`. */
function quasiPeriods(
	anchor: number,
	from: number,
	to: number,
	frequency: number,
	ctx: CallContext,
) {
	const months = 12 / frequency;
	const { year, month, day } = serialToYmd(anchor, ctx.date1904);
	const monthEnd = day === daysInMonth(year, month);
	const at = (k: number) => addMonths(anchor, k * months, ctx, monthEnd);
	let back = 0;
	while (at(back) > from) back -= 1;
	const dates: number[] = [];
	for (let k = back; ; k++) {
		dates.push(at(k));
		if (at(k) >= to) break;
	}
	return dates;
}

function accruedInterest(args: Value[], ctx: CallContext): number {
	const issue = serialArg(args[0]);
	const first = serialArg(args[1]);
	const settlement = serialArg(args[2]);
	const [rate, par] = [num(args[3]), optNum(args, 4, 1000)];
	const frequency = int(args[5]);
	const basis = basisOf(args, 6);
	const fromIssue = args.length > 7 && args[7] !== null ? bool(args[7]) : true;
	if (rate <= 0 || par <= 0 || ![1, 2, 4].includes(frequency) || issue >= settlement) fail(ERR.NUM);
	const start = fromIssue ? issue : Math.max(issue, first);
	if (start >= settlement) return 0;
	const dates = quasiPeriods(first, start, settlement, frequency, ctx);
	const count = (a: number, b: number) =>
		basis === 0 || basis === 4 ? days360(a, b, basis === 4, ctx) : b - a;
	let sum = 0;
	for (let i = 0; i + 1 < dates.length; i++) {
		const [p, q] = [dates[i] as number, dates[i + 1] as number];
		const a = count(Math.max(start, p), Math.min(settlement, q));
		if (a <= 0) continue;
		const length = basis === 1 ? q - p : basis === 3 ? 365 / frequency : 360 / frequency;
		sum += a / length;
	}
	return ((par * rate) / frequency) * sum;
}

const BOND_ARGS = (name: string, third: string, fourth: string, fifth: string) =>
	`${name}(settlement, maturity, ${third}, ${fourth}, ${fifth}, frequency, [basis])`;

export const BOND_FUNCTIONS: FunctionSpec[] = [
	spec(
		'PRICE',
		C,
		BOND_ARGS('PRICE', 'rate', 'yld', 'redemption'),
		'The price per 100 face value of a security that pays periodic interest.',
		6,
		7,
		(args, ctx) => {
			const bond = bondOf(args, ctx, 4, 6);
			const yld = num(args[3]);
			if (bond.rate < 0 || yld < 0 || bond.redemption <= 0) fail(ERR.NUM);
			return pricePerHundred(bond.facts, bond.rate, yld, bond.redemption, bond.frequency);
		},
	),
	spec(
		'YIELD',
		C,
		BOND_ARGS('YIELD', 'rate', 'pr', 'redemption'),
		'The yield of a security that pays periodic interest.',
		6,
		7,
		(args, ctx) => {
			const bond = bondOf(args, ctx, 4, 6);
			const price = num(args[3]);
			if (bond.rate < 0 || price <= 0 || bond.redemption <= 0) fail(ERR.NUM);
			const f = (y: number) =>
				pricePerHundred(bond.facts, bond.rate, y, bond.redemption, bond.frequency) - price;
			return solve(f, bond.rate || 0.05);
		},
	),
	...(['DURATION', 'MDURATION'] as const).map((name) =>
		spec(
			name,
			C,
			`${name}(settlement, maturity, coupon, yld, frequency, [basis])`,
			name === 'DURATION'
				? 'The Macaulay duration of a security with periodic interest.'
				: 'The modified Macaulay duration of a security.',
			5,
			6,
			(args, ctx) => {
				const frequency = int(args[4]);
				const facts = couponFacts(
					serialArg(args[0]),
					serialArg(args[1]),
					frequency,
					basisOf(args, 5),
					ctx,
				);
				const [rate, yld] = [num(args[2]), num(args[3])];
				if (rate < 0 || yld < 0) fail(ERR.NUM);
				const macaulay = duration({ facts, rate, redemption: 100, frequency }, yld);
				return name === 'DURATION' ? macaulay : macaulay / (1 + yld / frequency);
			},
		),
	),
	spec(
		'ACCRINT',
		C,
		'ACCRINT(issue, first_interest, settlement, rate, par, frequency, [basis], [calc_method])',
		'The accrued interest of a security that pays periodic interest.',
		6,
		8,
		(args, ctx) => accruedInterest(args, ctx),
	),
	spec(
		'ACCRINTM',
		C,
		'ACCRINTM(issue, settlement, rate, [par], [basis])',
		'The accrued interest of a security that pays interest at maturity.',
		3,
		5,
		(args, ctx) => {
			const [issue, settlement] = [serialArg(args[0]), serialArg(args[1])];
			const [rate, par] = [num(args[2]), optNum(args, 3, 1000)];
			if (rate <= 0 || par <= 0 || issue >= settlement) fail(ERR.NUM);
			return par * rate * yearFrac(issue, settlement, basisOf(args, 4), ctx);
		},
	),
];
