import type { Token } from './types.js';

const MS_PER_DAY = 86_400_000;
const EPOCH_1900 = Date.UTC(1899, 11, 30);
const EPOCH_1904 = Date.UTC(1904, 0, 1);
/** Largest serial Excel displays as a date (9999-12-31). */
export const MAX_DATE_SERIAL = 2_958_465;

/**
 * The date a serial number stands for, with the wall-clock time in the UTC fields. In the 1900
 * system serial 60 is Excel's fictitious 1900-02-29, which JavaScript cannot represent: it maps to
 * 1900-03-01 (as does serial 61); serials 1-59 follow the real calendar from 1900-01-01.
 */
export function serialToDate(serial: number, date1904 = false): Date {
	if (date1904) return new Date(EPOCH_1904 + Math.round(serial * MS_PER_DAY));
	const adjusted = serial < 61 ? serial + 1 : serial;
	return new Date(EPOCH_1900 + Math.round(adjusted * MS_PER_DAY));
}

/** The serial number of a date (its UTC fields are read as the wall-clock time). */
export function dateToSerial(date: Date, date1904 = false): number {
	const ms = date.getTime();
	if (date1904) return (ms - EPOCH_1904) / MS_PER_DAY;
	const serial = (ms - EPOCH_1900) / MS_PER_DAY;
	return serial < 61 ? serial - 1 : serial;
}

export interface DateParts {
	year: number;
	/** 1-12. */
	month: number;
	/** 0-31 (0 only for serial 0 in the 1900 system, shown as 1900-01-00). */
	day: number;
	/** 0 Sunday .. 6 Saturday. */
	weekday: number;
	hours: number;
	minutes: number;
	seconds: number;
	/** Fraction of a second, in units of 10^-precision. */
	subsec: number;
	/** Whole seconds elapsed since serial 0 (for `[h]`, `[m]`, `[s]`). */
	totalSeconds: number;
}

/** Calendar date for a whole day count. */
function calendar(days: number, date1904: boolean): { year: number; month: number; day: number } {
	if (!date1904) {
		if (days === 0) return { year: 1900, month: 1, day: 0 };
		if (days === 60) return { year: 1900, month: 2, day: 29 };
	}
	const d = new Date(
		(date1904 ? EPOCH_1904 : EPOCH_1900) + (date1904 || days > 60 ? days : days + 1) * MS_PER_DAY,
	);
	return { year: d.getUTCFullYear(), month: d.getUTCMonth() + 1, day: d.getUTCDate() };
}

/** Splits a serial into calendar and clock parts, rounding to `precision` second decimals. */
export function dateParts(serial: number, date1904: boolean, precision: number): DateParts {
	const unitsPerSecond = 10 ** precision;
	const unitsPerDay = 86_400 * unitsPerSecond;
	const total = Math.round(serial * unitsPerDay);
	const days = Math.floor(total / unitsPerDay);
	let rest = total - days * unitsPerDay;
	const subsec = rest % unitsPerSecond;
	rest = (rest - subsec) / unitsPerSecond;
	const { year, month, day } = calendar(days, date1904);
	return {
		year,
		month,
		day,
		weekday: (((days + (date1904 ? 5 : 6)) % 7) + 7) % 7,
		hours: Math.floor(rest / 3600),
		minutes: Math.floor(rest / 60) % 60,
		seconds: rest % 60,
		subsec,
		totalSeconds: Math.floor(total / unitsPerSecond),
	};
}

const MONTHS = [
	'January',
	'February',
	'March',
	'April',
	'May',
	'June',
	'July',
	'August',
	'September',
	'October',
	'November',
	'December',
];
const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

const two = (n: number): string => String(n).padStart(2, '0');

/** Seconds precision a date section needs (the longest `.0`, `.00`, `.000`). */
export function datePrecision(tokens: readonly Token[]): number {
	let p = 0;
	for (const t of tokens) if (t.t === 'subsec') p = Math.max(p, Math.min(3, t.digits));
	return p;
}

/** Formats a non-negative serial with a date/time section; `undefined` when out of range. */
export function renderDate(
	serial: number,
	tokens: readonly Token[],
	date1904: boolean,
): string | undefined {
	if (!(serial >= 0) || serial >= MAX_DATE_SERIAL + 1) return undefined;
	const precision = datePrecision(tokens);
	const p = dateParts(serial, date1904, precision);
	if (p.year > 9999) return undefined;
	const twelve = tokens.some((t) => t.t === 'ampm');
	const hour12 = p.hours % 12 === 0 ? 12 : p.hours % 12;
	let text = '';
	for (const tok of tokens) {
		if (tok.t === 'lit') text += tok.v;
		else if (tok.t === 'ampm') {
			const pm = p.hours >= 12;
			text +=
				tok.kind === 'AM/PM'
					? pm
						? 'PM'
						: 'AM'
					: tok.kind === 'A/P'
						? pm
							? 'P'
							: 'A'
						: pm
							? 'p'
							: 'a';
		} else if (tok.t === 'subsec') {
			const digits = String(p.subsec).padStart(precision, '0');
			text += `.${digits.slice(0, tok.digits).padEnd(tok.digits, '0')}`;
		} else if (tok.t === 'elapsed') {
			const amount =
				tok.unit === 'h'
					? Math.floor(p.totalSeconds / 3600)
					: tok.unit === 'm'
						? Math.floor(p.totalSeconds / 60)
						: p.totalSeconds;
			text += String(amount).padStart(tok.width, '0');
		} else if (tok.t === 'date') text += datePart(tok.part, p, twelve ? hour12 : p.hours);
	}
	return text;
}

function datePart(part: string, p: DateParts, hour: number): string {
	switch (part) {
		case 'y':
			return two(p.year % 100);
		case 'yyyy':
		case 'e':
			return String(p.year);
		case 'g':
			return '';
		case 'b':
			return two((p.year + 543) % 100);
		case 'bbbb':
			return String(p.year + 543);
		case 'm':
			return String(p.month);
		case 'mm':
			return two(p.month);
		case 'mmm':
			return (MONTHS[p.month - 1] ?? '').slice(0, 3);
		case 'mmmm':
			return MONTHS[p.month - 1] ?? '';
		case 'mmmmm':
			return (MONTHS[p.month - 1] ?? '').charAt(0);
		case 'd':
			return String(p.day);
		case 'dd':
			return two(p.day);
		case 'ddd':
			return (DAYS[p.weekday] ?? '').slice(0, 3);
		case 'dddd':
			return DAYS[p.weekday] ?? '';
		case 'h':
			return String(hour);
		case 'hh':
			return two(hour);
		case 'min':
			return String(p.minutes);
		case 'mmin':
			return two(p.minutes);
		case 's':
			return String(p.seconds);
		default:
			return two(p.seconds);
	}
}
