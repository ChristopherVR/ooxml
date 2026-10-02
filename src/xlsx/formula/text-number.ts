// Text-to-number coercion (`="10%"+0`, `="$1,000"*2`, `="2020-01-15"+1`) and Excel's
// number-to-text conversion (`=""&1/3` is `0.333333333333333`).
import { currentDate1904, ymdToSerial } from './date-serial.js';

const MONTHS = [
	'jan',
	'feb',
	'mar',
	'apr',
	'may',
	'jun',
	'jul',
	'aug',
	'sep',
	'oct',
	'nov',
	'dec',
] as const;

function monthFromName(name: string): number | undefined {
	const lower = name.toLowerCase();
	if (lower.length < 3) return undefined;
	const index = MONTHS.findIndex((m) => lower.startsWith(m));
	if (index < 0) return undefined;
	const full = new Date(Date.UTC(2000, index, 1))
		.toLocaleString('en-US', { month: 'long', timeZone: 'UTC' })
		.toLowerCase();
	return lower.length === 3 || full.startsWith(lower) || lower === 'sept' ? index + 1 : undefined;
}

const NUMBER_RE = /^(\d{1,3}(?:,\d{3})+|\d*)(\.\d*)?(?:e([+-]?\d+))?$/i;
const TIME_RE = /^(\d{1,4}):(\d{1,2})(?::(\d{1,2}(?:\.\d*)?))?\s*(am|pm|a|p)?$/i;

function parsePlainNumber(text: string): number | undefined {
	const match = NUMBER_RE.exec(text);
	if (!match) return undefined;
	const int = (match[1] ?? '').replace(/,/g, '');
	const frac = match[2] ?? '';
	if (int === '' && frac.length < 2) return undefined;
	const value = Number(`${int || '0'}${frac === '.' ? '' : frac}${match[3] ? `e${match[3]}` : ''}`);
	return Number.isFinite(value) ? value : undefined;
}

function parseTime(text: string): number | undefined {
	const match = TIME_RE.exec(text);
	if (!match) return undefined;
	let hours = Number(match[1]);
	const minutes = Number(match[2]);
	const seconds = match[3] ? Number(match[3]) : 0;
	const meridiem = match[4]?.toLowerCase();
	if (minutes > 59 || seconds >= 60) return undefined;
	if (meridiem) {
		if (hours < 1 || hours > 12) return undefined;
		if (hours === 12) hours = 0;
		if (meridiem.startsWith('p')) hours += 12;
	}
	return (hours * 3600 + minutes * 60 + seconds) / 86_400;
}

function serialOrUndefined(year: number, month: number, day: number): number | undefined {
	if (month < 1 || month > 12 || day < 1 || day > 31) return undefined;
	const check = new Date(Date.UTC(year, month - 1, day));
	if (check.getUTCMonth() !== month - 1 && !(year === 1900 && month === 2 && day === 29)) {
		return undefined;
	}
	return ymdToSerial(year, month, day, currentDate1904());
}

const fullYear = (text: string): number => {
	const year = Number(text);
	if (text.length > 2) return year;
	return year < 30 ? 2000 + year : 1900 + year;
};

function parseDate(text: string): number | undefined {
	const thisYear = new Date().getFullYear();
	let m = /^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/.exec(text);
	if (m) return serialOrUndefined(Number(m[1]), Number(m[2]), Number(m[3]));
	m = /^(\d{1,2})\/(\d{1,2})(?:\/(\d{2}|\d{4}))?$/.exec(text);
	if (m) return serialOrUndefined(m[3] ? fullYear(m[3]) : thisYear, Number(m[1]), Number(m[2]));
	m = /^(\d{1,2})[-\s]([a-z]+)\.?(?:[-\s,]+(\d{2}|\d{4}))?$/i.exec(text);
	if (m) {
		const month = monthFromName(m[2] ?? '');
		if (month) return serialOrUndefined(m[3] ? fullYear(m[3]) : thisYear, month, Number(m[1]));
	}
	m = /^([a-z]+)\.?[-\s](\d{1,2})(?:,?\s*(\d{4}))?$/i.exec(text);
	if (m) {
		const month = monthFromName(m[1] ?? '');
		if (month) return serialOrUndefined(m[3] ? Number(m[3]) : thisYear, month, Number(m[2]));
	}
	m = /^([a-z]+)\.?[-\s](\d{4})$/i.exec(text);
	if (m) {
		const month = monthFromName(m[1] ?? '');
		if (month) return serialOrUndefined(Number(m[2]), month, 1);
	}
	return undefined;
}

/**
 * The number a text converts to in arithmetic, or `undefined` when Excel would report `#VALUE!`.
 * Accepts thousands separators, currency, percent, parentheses for negatives, fractions, and
 * US-style dates and times.
 */
export function parseNumberText(input: string): number | undefined {
	let text = input.trim();
	if (text === '') return undefined;
	let scale = 1;
	if (text.endsWith('%')) {
		scale = 0.01;
		text = text.slice(0, -1).trimEnd();
	}
	let sign = 1;
	if (text.startsWith('(') && text.endsWith(')')) {
		sign = -1;
		text = text.slice(1, -1).trim();
	}
	const signed = /^([+-]?)\$?([+-]?)/.exec(text);
	const signText = `${signed?.[1] ?? ''}${signed?.[2] ?? ''}`;
	if (signText.length <= 1) {
		const body = text.slice(signed?.[0].length ?? 0);
		const plain = parsePlainNumber(body);
		if (plain !== undefined) return (signText === '-' ? -1 : 1) * sign * plain * scale;
	}
	if (scale !== 1 || sign !== 1) return undefined;
	const fraction = /^([+-]?)(?:(\d+)\s+)?(\d+)\/(\d+)$/.exec(text);
	if (fraction && fraction[2] !== undefined) {
		const den = Number(fraction[4]);
		if (den === 0) return undefined;
		const value = Number(fraction[2]) + Number(fraction[3]) / den;
		return fraction[1] === '-' ? -value : value;
	}
	const time = parseTime(text);
	if (time !== undefined) return time;
	const date = parseDate(text);
	if (date !== undefined) return date;
	const space = text.lastIndexOf(' ');
	if (space > 0) {
		const datePart = parseDate(text.slice(0, space).trim());
		const timePart = parseTime(text.slice(space + 1).trim());
		if (datePart !== undefined && timePart !== undefined) return datePart + timePart;
		const amPm = /^(.*\S)\s+(\S+\s*(?:am|pm))$/i.exec(text);
		if (amPm) {
			const d = parseDate(amPm[1] ?? '');
			const t = parseTime(amPm[2] ?? '');
			if (d !== undefined && t !== undefined) return d + t;
		}
	}
	return undefined;
}

/** Rounds to 15 significant digits, the precision Excel displays and compares with. */
export function round15(value: number): number {
	if (value === 0 || !Number.isFinite(value)) return value;
	return Number(value.toPrecision(15));
}

/** Excel's General conversion of a number to text (as `&`, CONCAT and friends produce it). */
export function numberToText(value: number): string {
	if (value === 0) return '0';
	if (!Number.isFinite(value)) return '#NUM!';
	const rounded = round15(value);
	const negative = rounded < 0;
	const abs = Math.abs(rounded);
	const [mantissa = '0', expText = '0'] = abs.toExponential(14).split('e');
	const exponent = Number(expText);
	const digits = mantissa.replace('.', '').replace(/0+$/, '') || '0';
	let plain: string;
	if (exponent >= 0) {
		plain =
			digits.length <= exponent + 1
				? digits + '0'.repeat(exponent + 1 - digits.length)
				: `${digits.slice(0, exponent + 1)}.${digits.slice(exponent + 1)}`;
	} else {
		plain = `0.${'0'.repeat(-exponent - 1)}${digits}`;
	}
	if (plain.length <= 20) return negative ? `-${plain}` : plain;
	const mant = digits.length > 1 ? `${digits[0]}.${digits.slice(1)}` : digits;
	const exp = `${exponent < 0 ? '-' : '+'}${String(Math.abs(exponent)).padStart(2, '0')}`;
	return `${negative ? '-' : ''}${mant}E${exp}`;
}
