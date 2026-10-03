import { dateToSerial } from './date.js';

export interface DateInput {
	value: number;
	numFmt: string;
}

const MONTH_NAMES = [
	'january',
	'february',
	'march',
	'april',
	'may',
	'june',
	'july',
	'august',
	'september',
	'october',
	'november',
	'december',
];

function monthFromName(name: string): number | undefined {
	const lower = name.toLowerCase();
	if (lower.length < 3) return undefined;
	const index = MONTH_NAMES.findIndex(
		(m) => m.startsWith(lower) || (lower.length > 3 && lower === m),
	);
	return index < 0 ? undefined : index + 1;
}

/** Two-digit years: 00-29 are 20xx, 30-99 are 19xx (Excel's cut-over). */
function fullYear(text: string): number {
	const y = Number(text);
	if (text.length > 2) return y;
	return y < 30 ? 2000 + y : 1900 + y;
}

/** Serial of a calendar date, or `undefined` when it is invalid or before the epoch. */
export function serialOf(
	year: number,
	month: number,
	day: number,
	date1904: boolean,
): number | undefined {
	if (month < 1 || month > 12 || day < 1 || year > 9999) return undefined;
	if (!date1904 && year === 1900 && month === 2 && day === 29) return 60;
	const date = new Date(Date.UTC(2000, month - 1, day));
	date.setUTCFullYear(year);
	if (date.getUTCMonth() !== month - 1) return undefined;
	const serial = dateToSerial(date, date1904);
	const minimum = date1904 ? 0 : 1;
	return serial < minimum ? undefined : serial;
}

interface TimeInput {
	fraction: number;
	format: string;
}

const TIME_RE = /^(\d{1,5}):(\d{1,2})(?::(\d{1,2}))?(?:\.(\d+))?\s*(am|pm|a|p)?$/i;

/** `h:mm`, `h:mm:ss`, `h:mm AM`, `[h]` beyond 24 hours and `mm:ss.0`. */
export function parseTime(text: string): TimeInput | undefined {
	const m = TIME_RE.exec(text);
	if (!m) return undefined;
	const [, a = '0', b = '0', c, frac, ampm] = m;
	let hours = Number(a);
	let minutes = Number(b);
	let seconds = c === undefined ? 0 : Number(c);
	let format: string;
	if (frac !== undefined) {
		if (ampm) return undefined;
		const secondsPart = Number(`${c ?? b}.${frac}`);
		// Excel picks `mm:ss.0` for every time with fractional seconds, even `10:30:00.5`.
		format = 'mm:ss.0';
		if (c === undefined) {
			minutes = hours;
			hours = 0;
		}
		seconds = secondsPart;
	} else if (ampm) {
		if (hours > 12) return undefined;
		const pm = ampm.toLowerCase().startsWith('p');
		hours = (hours % 12) + (pm ? 12 : 0);
		format = c === undefined ? 'h:mm AM/PM' : 'h:mm:ss AM/PM';
	} else if (hours >= 24) format = '[h]:mm:ss';
	else format = c === undefined ? 'h:mm' : 'h:mm:ss';
	if (minutes > 59 || seconds >= 60) return undefined;
	return { fraction: (hours * 3600 + minutes * 60 + seconds) / 86_400, format };
}

interface DayInput {
	year: number;
	month: number;
	day: number;
	format: string;
}

const currentYear = (): number => new Date().getFullYear();

function parseDay(text: string): DayInput | undefined {
	let m = /^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/.exec(text);
	if (m)
		return { year: Number(m[1]), month: Number(m[2]), day: Number(m[3]), format: 'yyyy-mm-dd' };
	m = /^(\d{1,2})([-/])(\d{1,2})\2(\d{2}|\d{4})$/.exec(text);
	if (m)
		return {
			year: fullYear(m[4] ?? ''),
			month: Number(m[1]),
			day: Number(m[3]),
			format: 'm/d/yyyy',
		};
	m = /^(\d{1,2})\/(\d{1,2})$/.exec(text);
	if (m) return { year: currentYear(), month: Number(m[1]), day: Number(m[2]), format: 'd-mmm' };
	m = /^(\d{1,2})[- ]([a-z]+)\.?(?:[- ,]+(\d{2}|\d{4}))?$/i.exec(text);
	if (m) {
		const month = monthFromName(m[2] ?? '');
		if (!month) return undefined;
		if (m[3] === undefined)
			return { year: currentYear(), month, day: Number(m[1]), format: 'd-mmm' };
		return { year: fullYear(m[3]), month, day: Number(m[1]), format: 'd-mmm-yy' };
	}
	m = /^([a-z]+)\.?[- ](\d{1,2}),?[- ](\d{4})$/i.exec(text);
	if (m) {
		const month = monthFromName(m[1] ?? '');
		return month ? { year: Number(m[3]), month, day: Number(m[2]), format: 'd-mmm-yy' } : undefined;
	}
	m = /^([a-z]{3,})[- ]?(\d{1,2})$/i.exec(text);
	if (m) {
		// `Jan 5`, `Jan-24`: a day of the current year when it exists, otherwise a two-digit year
		// (`Jan-32` is January 1932, `Feb-30` February 1930), as Excel reads them.
		const month = monthFromName(m[1] ?? '');
		if (!month) return undefined;
		const n = Number(m[2]);
		const year = currentYear();
		if (n >= 1 && serialOf(year, month, n, false) !== undefined)
			return { year, month, day: n, format: 'd-mmm' };
		return { year: fullYear(m[2] ?? ''), month, day: 1, format: 'mmm-yy' };
	}
	m = /^([a-z]+)\.?[- ](\d{4})$/i.exec(text);
	if (m) {
		const month = monthFromName(m[1] ?? '');
		return month ? { year: Number(m[2]), month, day: 1, format: 'mmm-yy' } : undefined;
	}
	return undefined;
}

/** Parses typed dates, times and date-times the way Excel's en-US entry does. */
export function parseDateTimeInput(text: string, date1904: boolean): DateInput | undefined {
	const time = parseTime(text);
	if (time) return { value: time.fraction, numFmt: time.format };
	const split = /^(.+?)(?:\s+|T)(\d{1,2}:\d{1,2}(?::\d{1,2})?(?:\s*(?:am|pm|a|p))?)$/i.exec(text);
	const dayText = split ? (split[1] ?? '') : text;
	const day = parseDay(dayText);
	if (!day) return undefined;
	const serial = serialOf(day.year, day.month, day.day, date1904);
	if (serial === undefined) return undefined;
	if (!split) return { value: serial, numFmt: day.format };
	const t = parseTime(split[2] ?? '');
	if (!t || t.format.startsWith('[')) return undefined;
	const dayFormat = day.format === 'd-mmm-yy' || day.format === 'd-mmm' ? 'm/d/yyyy' : day.format;
	return { value: serial + t.fraction, numFmt: `${dayFormat} ${t.format}` };
}
