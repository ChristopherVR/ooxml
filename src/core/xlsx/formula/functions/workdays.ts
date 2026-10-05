import type { CallContext } from '../context.js';
import { weekdayOf } from '../date-serial.js';
import { ERR, fail, isError, type Value } from '../values.js';
import { num, scalar, spec } from './helpers.js';
import type { FunctionSpec } from './types.js';

const C = 'Date & Time';

/** Weekend days (0 = Sunday) for NETWORKDAYS.INTL / WORKDAY.INTL weekend arguments. */
function weekendDays(value: Value | undefined): Set<number> {
	const v = scalar(value);
	if (v === null || value === undefined) return new Set([0, 6]);
	if (typeof v === 'string') {
		if (!/^[01]{7}$/.test(v)) fail(ERR.VALUE);
		const days = new Set<number>();
		// The string starts with Monday; all seven ("1111111") leaves no working days (WORKDAY.INTL
		// rejects it, NETWORKDAYS.INTL counts 0).
		for (let i = 0; i < 7; i++) if (v[i] === '1') days.add((i + 1) % 7);
		return days;
	}
	const code = Math.trunc(num(v));
	if (code >= 1 && code <= 7) {
		const first = (code + 5) % 7; // 1 -> Sat+Sun, 2 -> Sun+Mon, ...
		return new Set([first, (first + 1) % 7]);
	}
	if (code >= 11 && code <= 17) return new Set([code - 11]);
	return fail(ERR.NUM);
}

function holidaySet(ctx: CallContext, value: Value | undefined): Set<number> {
	const out = new Set<number>();
	if (value === undefined || value === null) return out;
	ctx.forEach(value, (v) => {
		if (isError(v)) fail(v);
		if (typeof v === 'number') out.add(Math.floor(v));
		else if (v !== null) out.add(Math.floor(num(v)));
	});
	return out;
}

function networkDays(
	start: number,
	end: number,
	weekend: Set<number>,
	holidays: Set<number>,
	ctx: CallContext,
): number {
	const a = Math.floor(start);
	const b = Math.floor(end);
	const step = a <= b ? 1 : -1;
	let count = 0;
	for (let d = a; step > 0 ? d <= b : d >= b; d += step) {
		if (!weekend.has(weekdayOf(d, ctx.date1904)) && !holidays.has(d)) count++;
	}
	return count * step;
}

function workday(
	start: number,
	days: number,
	weekend: Set<number>,
	holidays: Set<number>,
	ctx: CallContext,
): number {
	let d = Math.floor(start);
	let remaining = Math.trunc(days);
	const step = remaining >= 0 ? 1 : -1;
	if (weekend.size >= 7) fail(ERR.VALUE);
	while (remaining !== 0) {
		d += step;
		if (d < 0) fail(ERR.NUM);
		if (!weekend.has(weekdayOf(d, ctx.date1904)) && !holidays.has(d)) remaining -= step;
	}
	return d;
}

export const WORKDAY_FUNCTIONS: FunctionSpec[] = [
	spec(
		'NETWORKDAYS',
		C,
		'NETWORKDAYS(start_date, end_date, [holidays])',
		'Whole working days between two dates.',
		2,
		3,
		(args, ctx) =>
			networkDays(num(args[0]), num(args[1]), new Set([0, 6]), holidaySet(ctx, args[2]), ctx),
		['value', 'value', 'any'],
	),
	spec(
		'NETWORKDAYS.INTL',
		C,
		'NETWORKDAYS.INTL(start_date, end_date, [weekend], [holidays])',
		'Working days between dates with custom weekends.',
		2,
		4,
		(args, ctx) =>
			networkDays(num(args[0]), num(args[1]), weekendDays(args[2]), holidaySet(ctx, args[3]), ctx),
		['value', 'value', 'value', 'any'],
	),
	spec(
		'WORKDAY',
		C,
		'WORKDAY(start_date, days, [holidays])',
		'The date a number of working days away.',
		2,
		3,
		(args, ctx) =>
			workday(num(args[0]), num(args[1]), new Set([0, 6]), holidaySet(ctx, args[2]), ctx),
		['value', 'value', 'any'],
	),
	spec(
		'WORKDAY.INTL',
		C,
		'WORKDAY.INTL(start_date, days, [weekend], [holidays])',
		'The date a number of working days away with custom weekends.',
		2,
		4,
		(args, ctx) =>
			workday(num(args[0]), num(args[1]), weekendDays(args[2]), holidaySet(ctx, args[3]), ctx),
		['value', 'value', 'value', 'any'],
	),
];
