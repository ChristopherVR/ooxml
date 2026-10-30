// `<input type="date">` values and the serial numbers a date filter stores.
import { dateToSerial, serialToDate } from 'ooxml-core/xlsx';

/** `YYYY-MM-DD` to a whole-day serial, or undefined for an empty or invalid value. */
export function dateInputToSerial(value: string, date1904: boolean): number | undefined {
	const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(value);
	if (!match) return undefined;
	const ms = Date.UTC(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
	return Number.isNaN(ms) ? undefined : Math.floor(dateToSerial(new Date(ms), date1904));
}

/** A stored serial as `YYYY-MM-DD` ('' when it is not a number). */
export function serialToDateInput(serial: string, date1904 = false): string {
	const n = Number(serial);
	if (serial.trim() === '' || !Number.isFinite(n)) return '';
	return serialToDate(n, date1904).toISOString().slice(0, 10);
}
