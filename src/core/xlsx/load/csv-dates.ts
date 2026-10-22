// Date columns of a CSV import (the Text Import Wizard's Date column format): the day, month and
// year are read in a stated order instead of en-US month first, optionally followed by a time
// and an ISO 8601 time-zone designator.
import { fullYear, parseTime, serialOf } from '../numfmt/input-date';

/** The order of day (D), month (M) and year (Y) in a date column, as the wizard offers them. */
export type CsvDateOrder = 'MDY' | 'DMY' | 'YMD' | 'MYD' | 'DYM' | 'YDM';

export interface OrderedDate {
	/** The date serial, with the time of day as its fraction. */
	value: number;
	numFmt: string;
	/** The time carried a UTC designator or offset and was converted to UTC. */
	zoned: boolean;
}

/** Three numeric parts with one separator (`04/03/2024`, `4-3-24`, `04.03.2024`, `4 3 2024`). */
const DAY_RE = /^(\d{1,4})([-/. ])(\d{1,4})\2(\d{1,4})$/;
/** Eight digits with no separator (`20240304` in a YMD column). */
const COMPACT_RE = /^\d{8}$/;
/** A day and a month in the current year (`04/03`), as typing reads `m/d`. */
const PAIR_RE = /^(\d{1,2})\/(\d{1,2})$/;
/** The fraction of a second in a time (`10:20:30.250`). */
const FRACTION_RE = /\.(\d+)/;
/** A time (seconds, fractions and AM/PM optional), then `Z` or `+hh[:mm]`; anchored at both ends. */
const TIME_RE =
	/^(\d{1,2}:\d{1,2}(?::\d{1,2}(?:\.\d+)?)?(?:\s*[ap]m?)?)\s*(Z|[+-]\d{2}(?::?\d{2})?)?$/i;
const SPACE_RE = /\s/;
/**
 * The longest field read as a date. No date and time needs more, and it caps the work of the
 * patterns, whatever padding a file holds.
 */
export const MAX_DATE_TEXT = 64;

/**
 * The date and the time of `text`, found in one pass: the time starts at the one or two hour
 * digits before the first `:` and follows a space or `T`. `undefined` when a `:` has no such time.
 */
function splitTime(text: string): [day: string, time: string | undefined] | undefined {
	const colon = text.indexOf(':');
	if (colon < 0) return [text, undefined];
	let start = colon;
	while (start > 0 && colon - start < 2) {
		const code = text.charCodeAt(start - 1);
		if (code < 48 || code > 57) break;
		start--;
	}
	const before = text[start - 1] ?? '';
	if (start === colon || !before) return undefined;
	if (before === 'T' || before === 't') return [text.slice(0, start - 1), text.slice(start)];
	return SPACE_RE.test(before) ? [text.slice(0, start).trimEnd(), text.slice(start)] : undefined;
}

/** Splits eight digits by `order`: four for the year, two for the day and the month. */
function compactParts(text: string, order: CsvDateOrder): string[] {
	const parts: string[] = [];
	let at = 0;
	for (const kind of order) {
		const size = kind === 'Y' ? 4 : 2;
		parts.push(text.slice(at, at + size));
		at += size;
	}
	return parts;
}

/** The day and month of `04/03` in `order`, with the current year in the year's place. */
function pairParts(pair: readonly string[], order: CsvDateOrder): string[] {
	let next = 0;
	return [...order].map((kind) =>
		kind === 'Y' ? String(new Date().getFullYear()) : (pair[next++] ?? ''),
	);
}

/** The three date parts of `text` in `order`'s positions (separated, compact or day and month). */
function dayParts(text: string, order: CsvDateOrder): string[] | undefined {
	const match = DAY_RE.exec(text);
	if (match) return [match[1] ?? '', match[3] ?? '', match[4] ?? ''];
	if (COMPACT_RE.test(text)) return compactParts(text, order);
	const pair = PAIR_RE.exec(text);
	return pair ? pairParts([pair[1] ?? '', pair[2] ?? ''], order) : undefined;
}

/** Year, month and day from parts in `order`; a one- or two-digit year uses Excel's cut-over. */
function inOrder(
	parts: readonly string[],
	order: CsvDateOrder,
): [year: number, month: number, day: number] | undefined {
	const date: [year: number, month: number, day: number] = [0, 0, 0];
	for (let i = 0; i < 3; i++) {
		const part = parts[i] ?? '';
		const kind = order[i];
		if (kind === 'Y') {
			if (part.length === 3 || part.length > 4) return undefined;
			date[0] = fullYear(part);
		} else if (part.length > 2) return undefined;
		else date[kind === 'M' ? 1 : 2] = Number(part);
	}
	return date;
}

/** Minutes east of UTC for `Z`, `+hh`, `+hhmm` or `+hh:mm`. */
function offsetMinutes(zone: string): number | undefined {
	if (zone === 'Z' || zone === 'z') return 0;
	const hours = Number(zone.slice(1, 3));
	const minutes = zone.length > 3 ? Number(zone.slice(-2)) : 0;
	if (hours > 23 || minutes > 59) return undefined;
	return (zone.startsWith('-') ? -1 : 1) * (hours * 60 + minutes);
}

/**
 * A date (and optional time) whose day, month and year appear in `order`, or `undefined` when
 * `text` is not one. The model holds no time zone, so a time with `Z` or an offset is converted
 * to UTC and reported through `zoned`. Values show as `yyyy-mm-dd` (plus the time), which reads
 * the same whatever the source order was.
 */
export function parseOrderedDate(
	text: string,
	order: CsvDateOrder,
	date1904: boolean,
	decimal = '.',
): OrderedDate | undefined {
	if (text.length > MAX_DATE_TEXT) return undefined;
	const split = splitTime(text);
	if (!split) return undefined;
	// Fractions of a second use the import's decimal separator (`10:20:30,5`), and only it.
	let clock = split[1];
	if (clock !== undefined && decimal !== '.')
		clock = clock.includes('.') ? '' : clock.replace(decimal, '.');
	const tail = clock === undefined ? undefined : TIME_RE.exec(clock);
	if (clock !== undefined && !tail) return undefined;
	const parts = dayParts(split[0], order);
	const date = parts && inOrder(parts, order);
	if (!date) return undefined;
	const serial = serialOf(date[0], date[1], date[2], date1904);
	if (serial === undefined) return undefined;
	if (!tail) return { value: serial, numFmt: 'yyyy-mm-dd', zoned: false };
	// Typing reads `25:00` (and `25:00:00.5`) as elapsed time; a time of day stays below 24 hours.
	const time = parseTime(tail[1] ?? '');
	if (!time || time.fraction >= 1) return undefined;
	const zone = tail[2];
	const offset = zone === undefined ? 0 : offsetMinutes(zone);
	if (offset === undefined) return undefined;
	const value = serial + time.fraction - offset / 1440;
	if (value < (date1904 ? 0 : 1)) return undefined;
	// Show as many fraction digits as the field has, up to milliseconds (Excel's `ss.000`).
	const digits = Math.min(3, FRACTION_RE.exec(tail[1] ?? '')?.[1]?.length ?? 0);
	const format = digits ? `h:mm:ss.${'0'.repeat(digits)}` : time.format;
	return { value, numFmt: `yyyy-mm-dd ${format}`, zoned: zone !== undefined };
}
