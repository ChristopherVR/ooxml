// Calendar arithmetic behind the date functions.
import type { CallContext } from '../context.js';
import {
	daysInMonth,
	isLeapYear,
	monthStartSerial,
	serialToYmd,
	weekdayOf,
	ymdToSerial,
} from '../date-serial.js';
import { ERR, fail, type Value } from '../values.js';
import { num } from './helpers.js';

const MAX_SERIAL = 2_958_465;

/** A date serial argument: numbers and date text, truncated; negative or too large is #NUM!. */
export function serialArg(value: Value | undefined): number {
	const n = num(value);
	if (n < 0 || n > MAX_SERIAL + 0.99999999) fail(ERR.NUM);
	return n;
}

export function dateSerial(year: number, month: number, day: number, ctx: CallContext): number {
	let y = Math.trunc(year);
	if (y < 0 || y >= 10_000) fail(ERR.NUM);
	if (y < 1900) y += 1900;
	const serial = ymdToSerial(y, Math.trunc(month), Math.trunc(day), ctx.date1904);
	return serial === undefined || serial > MAX_SERIAL ? fail(ERR.NUM) : serial;
}

export function addMonths(
	serial: number,
	months: number,
	ctx: CallContext,
	endOfMonth: boolean,
): number {
	const { year, month, day } = serialToYmd(serialArg(serial), ctx.date1904);
	const total = year * 12 + (month - 1) + Math.trunc(months);
	const y = Math.floor(total / 12);
	const m = (total % 12) + 1;
	const d = endOfMonth ? daysInMonth(y, m) : Math.min(day, daysInMonth(y, m));
	const out = ymdToSerial(y, m, d, ctx.date1904);
	return out === undefined ? fail(ERR.NUM) : out;
}

/** WEEKDAY's numbering: returns 1-based (or 0-based for type 3) day number. */
export function weekday(serial: number, type: number, ctx: CallContext): number {
	const dow = weekdayOf(serial, ctx.date1904);
	const startOf: Record<number, number> = {
		1: 0,
		2: 1,
		11: 1,
		12: 2,
		13: 3,
		14: 4,
		15: 5,
		16: 6,
		17: 0,
	};
	if (type === 3) return (dow + 6) % 7;
	const start = startOf[type];
	if (start === undefined) fail(ERR.NUM);
	return ((dow - start + 7) % 7) + 1;
}

/**
 * ISO week number on Excel's calendar, whose weekdays before 1900-03-01 follow the phantom
 * 1900-02-29 (serial 1 is a Sunday, so ISOWEEKNUM(1) is 52).
 */
export function isoWeek(serial: number, ctx: CallContext): number {
	const days = Math.floor(serial);
	const thursday = days - ((weekdayOf(days, ctx.date1904) + 6) % 7) + 3;
	const year = !ctx.date1904 && thursday < 1 ? 1899 : serialToYmd(thursday, ctx.date1904).year;
	return Math.floor((thursday - monthStartSerial(year, 0, ctx.date1904)) / 7) + 1;
}

export function weekNum(serial: number, type: number, ctx: CallContext): number {
	if (type === 21) return isoWeek(serial, ctx);
	const startOf: Record<number, number> = {
		1: 0,
		2: 1,
		11: 1,
		12: 2,
		13: 3,
		14: 4,
		15: 5,
		16: 6,
		17: 0,
	};
	const start = startOf[type];
	if (start === undefined) fail(ERR.NUM);
	const { year } = serialToYmd(serial, ctx.date1904);
	const jan1 = ymdToSerial(year, 1, 1, ctx.date1904) ?? 0;
	const offset = (weekdayOf(jan1, ctx.date1904) - start + 7) % 7;
	return Math.floor((Math.floor(serial) - jan1 + offset) / 7) + 1;
}

export function datedif(start: number, end: number, unit: string, ctx: CallContext): number {
	if (start > end) fail(ERR.NUM);
	const a = serialToYmd(start, ctx.date1904);
	const b = serialToYmd(end, ctx.date1904);
	let months = (b.year - a.year) * 12 + (b.month - a.month);
	if (b.day < a.day) months--;
	switch (unit.toUpperCase()) {
		case 'Y':
			return Math.floor(months / 12);
		case 'M':
			return months;
		case 'D':
			return Math.floor(end) - Math.floor(start);
		case 'YM':
			return months % 12;
		case 'MD': {
			if (b.day >= a.day) return b.day - a.day;
			const pm = b.month === 1 ? 12 : b.month - 1;
			const py = b.month === 1 ? b.year - 1 : b.year;
			return daysInMonth(py, pm) - a.day + b.day;
		}
		case 'YD': {
			// The anniversary overflows like DATE (a 29 February start falls on 1 March).
			let y = b.year;
			let anniversary = ymdToSerial(y, a.month, a.day, ctx.date1904) ?? 0;
			if (anniversary > Math.floor(end)) {
				y--;
				anniversary = ymdToSerial(y, a.month, a.day, ctx.date1904) ?? 0;
			}
			return Math.floor(end) - anniversary;
		}
		default:
			return fail(ERR.NUM);
	}
}

export function days360(start: number, end: number, european: boolean, ctx: CallContext): number {
	const a = serialToYmd(start, ctx.date1904);
	const b = serialToYmd(end, ctx.date1904);
	let d1 = a.day;
	let d2 = b.day;
	if (european) {
		if (d1 === 31) d1 = 30;
		if (d2 === 31) d2 = 30;
	} else {
		const lastFeb = (y: number, m: number, d: number): boolean =>
			m === 2 && d === daysInMonth(y, 2);
		if (lastFeb(a.year, a.month, d1)) d1 = 30;
		if (d1 === 31) d1 = 30;
		if (d2 === 31 && d1 >= 30) d2 = 30;
	}
	return (b.year - a.year) * 360 + (b.month - a.month) * 30 + (d2 - d1);
}

/** YEARFRAC's US 30/360 day count, whose end-of-February rules differ from DAYS360's. */
function yearFrac30(a: { year: number; month: number; day: number }, b: typeof a): number {
	const lastFeb = (d: typeof a): boolean => d.month === 2 && d.day === daysInMonth(d.year, 2);
	let d1 = a.day;
	let d2 = b.day;
	if (d1 === 31 && d2 === 31) {
		d1 = 30;
		d2 = 30;
	} else if (d1 === 31) d1 = 30;
	else if (d1 === 30 && d2 === 31) d2 = 30;
	else if (lastFeb(a) && lastFeb(b)) {
		d1 = 30;
		d2 = 30;
	} else if (lastFeb(a)) d1 = 30;
	return ((b.year - a.year) * 360 + (b.month - a.month) * 30 + (d2 - d1)) / 360;
}

export function yearFrac(s: number, e: number, basis: number, ctx: CallContext): number {
	let start = Math.floor(s);
	let end = Math.floor(e);
	if (start > end) [start, end] = [end, start];
	const a = serialToYmd(start, ctx.date1904);
	const b = serialToYmd(end, ctx.date1904);
	switch (basis) {
		case 0:
			return yearFrac30(a, b);
		case 1: {
			if (a.year === b.year) return (end - start) / (isLeapYear(a.year) ? 366 : 365);
			const oneYear =
				b.year === a.year + 1 && (b.month < a.month || (b.month === a.month && b.day <= a.day));
			if (oneYear) {
				const feb29 =
					(isLeapYear(a.year) && a.month < 3) ||
					(isLeapYear(b.year) && (b.month > 2 || (b.month === 2 && b.day === 29)));
				return (end - start) / (feb29 ? 366 : 365);
			}
			let total = 0;
			for (let y = a.year; y <= b.year; y++) total += isLeapYear(y) ? 366 : 365;
			return (end - start) / (total / (b.year - a.year + 1));
		}
		case 2:
			return (end - start) / 360;
		case 3:
			return (end - start) / 365;
		case 4:
			return days360(start, end, true, ctx) / 360;
		default:
			return fail(ERR.NUM);
	}
}
