// Date-serial helpers for the formula engine. They stay separate from numfmt's public
// serialToDate/dateToSerial on purpose: functions need serial 60 as 1900-02-29 and serial 0 as
// 1900-01-00 (DAY, MONTH, YEAR, WEEKDAY), which a JavaScript Date cannot represent.
// Serial numbers count days from the workbook epoch. In the 1900 system Excel keeps Lotus 1-2-3's
// leap-year bug: serial 60 is the non-existent 1900-02-29, serial 0 is "1900-01-00".

const DAY_MS = 86_400_000;
/** UTC milliseconds of 1899-12-30, the day before serial 1 counting the phantom leap day. */
const EPOCH_1900 = Date.UTC(1899, 11, 30);
const EPOCH_1904 = Date.UTC(1904, 0, 1);

export interface Ymd {
	year: number;
	month: number;
	day: number;
}

let date1904Mode = false;

/** The date system used by coercions inside the current evaluation. */
export const currentDate1904 = (): boolean => date1904Mode;

/** Runs `fn` with the given date system active for text-to-date coercion. */
export function withDateSystem<T>(date1904: boolean, fn: () => T): T {
	const previous = date1904Mode;
	date1904Mode = date1904;
	try {
		return fn();
	} finally {
		date1904Mode = previous;
	}
}

/** The calendar date of the integer part of `serial`. */
export function serialToYmd(serial: number, date1904 = date1904Mode): Ymd {
	const days = Math.floor(serial);
	if (date1904) {
		const d = new Date(EPOCH_1904 + days * DAY_MS);
		return { year: d.getUTCFullYear(), month: d.getUTCMonth() + 1, day: d.getUTCDate() };
	}
	if (days === 0) return { year: 1900, month: 1, day: 0 };
	if (days === 60) return { year: 1900, month: 2, day: 29 };
	const d = new Date(EPOCH_1900 + (days < 60 ? days + 1 : days) * DAY_MS);
	return { year: d.getUTCFullYear(), month: d.getUTCMonth() + 1, day: d.getUTCDate() };
}

/**
 * The serial of a calendar date, normalising month and day overflow like Excel's DATE (month 13
 * is January of the next year, day 0 is the last day of the previous month). Returns `undefined`
 * for dates before the epoch or after 9999-12-31.
 */
export function ymdToSerial(
	year: number,
	month: number,
	day: number,
	date1904 = date1904Mode,
): number | undefined {
	const utc = Date.UTC(year, month - 1, 1) + (day - 1) * DAY_MS;
	if (!Number.isFinite(utc)) return undefined;
	if (new Date(utc).getUTCFullYear() > 9999) return undefined;
	if (date1904) {
		const serial = Math.round((utc - EPOCH_1904) / DAY_MS);
		return serial < 0 ? undefined : serial;
	}
	if (year === 1900 && month === 2 && day === 29) return 60;
	let serial = Math.round((utc - EPOCH_1900) / DAY_MS);
	if (serial < 61) serial -= 1;
	return serial < 0 ? undefined : serial;
}

/** Day of the week of a serial, 0 = Sunday (honours the 1900 bug: serial 1 is a Sunday). */
export function weekdayOf(serial: number, date1904 = date1904Mode): number {
	const days = Math.floor(serial);
	if (date1904) return (((days + 5) % 7) + 7) % 7;
	// Serial 1 (1900-01-01) is reported as a Sunday by Excel (really a Monday).
	return (((days - 1) % 7) + 7) % 7;
}

/** Whole seconds in the fractional part of a serial, rounded like Excel's HOUR/MINUTE/SECOND. */
export function serialSeconds(serial: number): number {
	const fraction = serial - Math.floor(serial);
	return Math.round(fraction * 86_400) % 86_400;
}

export const isLeapYear = (year: number): boolean =>
	(year % 4 === 0 && year % 100 !== 0) || year % 400 === 0;

export function daysInMonth(year: number, month: number): number {
	return new Date(Date.UTC(year, month, 0)).getUTCDate();
}

/** The serial for a JavaScript Date's local wall-clock time (for NOW and TODAY). */
export function localDateToSerial(date: Date, date1904 = date1904Mode): number {
	const utc = Date.UTC(
		date.getFullYear(),
		date.getMonth(),
		date.getDate(),
		date.getHours(),
		date.getMinutes(),
		date.getSeconds(),
		date.getMilliseconds(),
	);
	const dayStart = Date.UTC(date.getFullYear(), date.getMonth(), date.getDate());
	const base = ymdToSerial(date.getFullYear(), date.getMonth() + 1, date.getDate(), date1904) ?? 0;
	return base + (utc - dayStart) / DAY_MS;
}
