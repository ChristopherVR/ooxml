import type { DateGroupItem, DynamicFilterType, TimePeriod } from './model';
import { dateToSerial, serialToDate } from './numfmt/date';
import { inTimePeriod } from './time-period';

/** Today's date serial in the local time zone. */
export function todaySerial(date1904: boolean): number {
	const now = new Date();
	return Math.floor(
		dateToSerial(new Date(Date.UTC(now.getFullYear(), now.getMonth(), now.getDate())), date1904),
	);
}

const PERIODS: Partial<Record<DynamicFilterType, TimePeriod>> = {
	today: 'today',
	yesterday: 'yesterday',
	tomorrow: 'tomorrow',
	thisWeek: 'thisWeek',
	lastWeek: 'lastWeek',
	nextWeek: 'nextWeek',
	thisMonth: 'thisMonth',
	lastMonth: 'lastMonth',
	nextMonth: 'nextMonth',
};

/** Whether the date serial `value` satisfies a date-based dynamic filter relative to `today`. */
export function inDynamicPeriod(
	type: DynamicFilterType,
	value: number,
	today: number,
	date1904: boolean,
): boolean {
	const period = PERIODS[type];
	if (period) return inTimePeriod(period, value, today, date1904);
	const d = serialToDate(Math.floor(value), date1904);
	const now = serialToDate(today, date1904);
	const quarter = (date: Date) => date.getUTCFullYear() * 4 + Math.floor(date.getUTCMonth() / 3);
	switch (type) {
		case 'thisQuarter':
			return quarter(d) === quarter(now);
		case 'lastQuarter':
			return quarter(d) === quarter(now) - 1;
		case 'nextQuarter':
			return quarter(d) === quarter(now) + 1;
		case 'thisYear':
			return d.getUTCFullYear() === now.getUTCFullYear();
		case 'lastYear':
			return d.getUTCFullYear() === now.getUTCFullYear() - 1;
		case 'nextYear':
			return d.getUTCFullYear() === now.getUTCFullYear() + 1;
		case 'yearToDate':
			return d.getUTCFullYear() === now.getUTCFullYear() && Math.floor(value) <= today;
		default:
	}
	const q = /^Q([1-4])$/.exec(type);
	if (q) return Math.floor(d.getUTCMonth() / 3) + 1 === Number(q[1]);
	const m = /^M(\d{1,2})$/.exec(type);
	if (m) return d.getUTCMonth() + 1 === Number(m[1]);
	return false;
}

/** Whether a date serial lies inside a `dateGroupItem` (down to the group's own level). */
export function inDateGroup(item: DateGroupItem, serial: number, date1904: boolean): boolean {
	const d = serialToDate(serial, date1904);
	const fields: [number | undefined, number][] = [
		[item.year, d.getUTCFullYear()],
		[item.month, d.getUTCMonth() + 1],
		[item.day, d.getUTCDate()],
		[item.hour, d.getUTCHours()],
		[item.minute, d.getUTCMinutes()],
		[item.second, d.getUTCSeconds()],
	];
	const depth = ['year', 'month', 'day', 'hour', 'minute', 'second'].indexOf(item.grouping);
	return fields.slice(0, depth + 1).every(([want, have]) => want === undefined || want === have);
}
