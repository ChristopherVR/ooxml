// Coupon schedules for bonds (the COUP* functions) and the day counts the bond prices share.
import type { CallContext } from '../context';
import { daysInMonth, serialToYmd } from '../date-serial';
import { ERR, fail, type Value } from '../values';
import { addMonths, days360, serialArg } from './date-calc';
import { int, spec } from './helpers';
import type { FunctionSpec } from './types';

const C = 'Financial';

export interface CouponFacts {
	/** The last coupon date on or before settlement. */
	previous: number;
	/** The first coupon date after settlement. */
	next: number;
	/** Coupons payable between settlement and maturity. */
	count: number;
	/** Days from the previous coupon to settlement. */
	daysSincePrevious: number;
	/** Days in the coupon period holding settlement. */
	daysInPeriod: number;
	/** Days from settlement to the next coupon. */
	daysToNext: number;
}

export function checkSecurity(
	settlement: number,
	maturity: number,
	frequency: number,
	basis: number,
): void {
	if (![1, 2, 4].includes(frequency) || basis < 0 || basis > 4 || settlement >= maturity)
		fail(ERR.NUM);
}

/** Coupon dates run back from maturity, on month ends when maturity is one (Excel's rule). */
export function couponFacts(
	settlement: number,
	maturity: number,
	frequency: number,
	basis: number,
	ctx: CallContext,
): CouponFacts {
	const [s, m] = [Math.floor(settlement), Math.floor(maturity)];
	checkSecurity(s, m, frequency, basis);
	const months = 12 / frequency;
	const { year, month, day } = serialToYmd(m, ctx.date1904);
	const monthEnd = day === daysInMonth(year, month);
	let count = 0;
	let next = m;
	let previous = addMonths(m, -months, ctx, monthEnd);
	while (previous > s) {
		count += 1;
		next = previous;
		previous = addMonths(m, -months * (count + 1), ctx, monthEnd);
	}
	count += 1;
	const dayCount = (a: number, b: number, european: boolean) =>
		basis === 0 || basis === 4 ? days360(a, b, basis === 4 || european, ctx) : b - a;
	const daysSincePrevious = dayCount(previous, s, false);
	const daysInPeriod =
		basis === 1 ? next - previous : basis === 3 ? 365 / frequency : 360 / frequency;
	const daysToNext =
		basis === 0
			? daysInPeriod - daysSincePrevious
			: basis === 4
				? days360(s, next, true, ctx)
				: next - s;
	return { previous, next, count, daysSincePrevious, daysInPeriod, daysToNext };
}

const security = (
	name: string,
	description: string,
	pick: (facts: CouponFacts) => number,
): FunctionSpec =>
	spec(
		name,
		C,
		`${name}(settlement, maturity, frequency, [basis])`,
		description,
		3,
		4,
		(args: Value[], ctx) =>
			pick(
				couponFacts(
					serialArg(args[0]),
					serialArg(args[1]),
					int(args[2]),
					args.length > 3 && args[3] !== null ? int(args[3]) : 0,
					ctx,
				),
			),
	);

export const COUPON_FUNCTIONS: FunctionSpec[] = [
	security('COUPPCD', 'The coupon date before the settlement date.', (f) => f.previous),
	security('COUPNCD', 'The coupon date after the settlement date.', (f) => f.next),
	security('COUPNUM', 'The number of coupons between settlement and maturity.', (f) => f.count),
	security(
		'COUPDAYBS',
		'The days from the start of the coupon period to settlement.',
		(f) => f.daysSincePrevious,
	),
	security(
		'COUPDAYS',
		'The days in the coupon period that contains the settlement date.',
		(f) => f.daysInPeriod,
	),
	security('COUPDAYSNC', 'The days from settlement to the next coupon date.', (f) => f.daysToNext),
];
