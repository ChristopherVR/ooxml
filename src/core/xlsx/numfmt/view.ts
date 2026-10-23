// What a reader reports about a value's number format besides the value itself: the format code
// and, for a number the format shows as a date or time, that date in ISO 8601, so a caller can
// read dates without knowing Excel's serial numbers. The automation reads and the row reader
// report the same.
import type { CellValue } from '../model';
import { MAX_DATE_SERIAL, serialToDate } from './date';
import { dateTimeParts } from './format';

/**
 * ISO 8601 without a time zone for a date or time serial: a date when it is whole and the format
 * shows no time, a time of day when there is no calendar day to name (below 1 in the 1900 system,
 * whose day 0 is fictitious, or a time-only format in the 1904 one), both otherwise, with
 * milliseconds when there are any. `undefined` before the epoch, after 9999 and for the fictitious
 * 1900-02-29 (serial 60).
 */
function isoOf(
	serial: number,
	date1904: boolean,
	shows: { date: boolean; time: boolean },
): string | undefined {
	// `!(serial >= 0)` also refuses NaN; past 9999 a date may not exist in JavaScript at all.
	if (!(serial >= 0) || serial > MAX_DATE_SERIAL + 1) return undefined;
	if (!date1904 && serial >= 60 && serial < 61) return undefined;
	const date = serialToDate(serial, date1904);
	if (date.getUTCFullYear() > 9999) return undefined;
	const iso = date.toISOString();
	const time = iso.slice(11, iso.endsWith('.000Z') ? 19 : 23);
	if (serial < 1 && (!date1904 || !shows.date)) return time;
	return Number.isInteger(serial) && !shows.time ? iso.slice(0, 10) : `${iso.slice(0, 10)}T${time}`;
}

export interface NumberFormatView {
	/** The number format code, absent for General. */
	numberFormat?: string;
	/**
	 * The value of a cell shown as a date or time, in ISO 8601 without a time zone; absent for
	 * numbers and durations (`[h]:mm`).
	 */
	iso?: string;
}

/** The number format and ISO date a cell with format `numFmt` and `value` reports. */
export function numberFormatView(
	numFmt: string,
	value: CellValue,
	date1904: boolean,
): NumberFormatView {
	if (numFmt === 'General') return {};
	const parts = typeof value === 'number' ? dateTimeParts(numFmt, value) : undefined;
	const iso = parts && typeof value === 'number' ? isoOf(value, date1904, parts) : undefined;
	return iso ? { numberFormat: numFmt, iso } : { numberFormat: numFmt };
}
