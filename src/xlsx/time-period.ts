// "A Date Occurring" conditional-format periods: the formula Excel stores with the rule and a
// direct test used by the layout evaluator.
import type { TimePeriod } from './model.js';
import { serialToDate } from './numfmt/date.js';

/** The formula Excel writes for a `timePeriod` rule whose range starts at `cell` (`A1`). */
export function timePeriodFormula(period: TimePeriod, cell: string): string {
	const day = `FLOOR(${cell},1)`;
	const down = `ROUNDDOWN(${cell},0)`;
	const month = (offset: string) =>
		`AND(MONTH(${cell})=MONTH(EDATE(TODAY(),${offset})),YEAR(${cell})=YEAR(EDATE(TODAY(),${offset})))`;
	switch (period) {
		case 'today':
			return `${day}=TODAY()`;
		case 'yesterday':
			return `${day}=TODAY()-1`;
		case 'tomorrow':
			return `${day}=TODAY()+1`;
		case 'last7Days':
			return `AND(TODAY()-${day}<=6,${day}<=TODAY())`;
		case 'thisWeek':
			return `AND(TODAY()-${down}<=WEEKDAY(TODAY())-1,${down}-TODAY()<=7-WEEKDAY(TODAY()))`;
		case 'lastWeek':
			return `AND(TODAY()-${down}>=(WEEKDAY(TODAY())),TODAY()-${down}<(WEEKDAY(TODAY())+7))`;
		case 'nextWeek':
			return `AND(${down}-TODAY()>(7-WEEKDAY(TODAY())),${down}-TODAY()<(15-WEEKDAY(TODAY())))`;
		case 'thisMonth':
			return `AND(MONTH(${cell})=MONTH(TODAY()),YEAR(${cell})=YEAR(TODAY()))`;
		case 'lastMonth':
			return month('0-1');
		case 'nextMonth':
			return month('0+1');
	}
}

/**
 * Whether the date serial `value` falls in `period` relative to the serial `today` (only the
 * whole-day parts count), with weeks running Sunday to Saturday as in Excel.
 */
export function inTimePeriod(
	period: TimePeriod,
	value: number,
	today: number,
	date1904 = false,
): boolean {
	const v = Math.floor(value);
	const t = Math.floor(today);
	const todayDate = serialToDate(t, date1904);
	const weekday = todayDate.getUTCDay() + 1;
	const monthOf = (serial: number) => {
		const d = serialToDate(serial, date1904);
		return d.getUTCFullYear() * 12 + d.getUTCMonth();
	};
	const current = todayDate.getUTCFullYear() * 12 + todayDate.getUTCMonth();
	switch (period) {
		case 'today':
			return v === t;
		case 'yesterday':
			return v === t - 1;
		case 'tomorrow':
			return v === t + 1;
		case 'last7Days':
			return t - v <= 6 && v <= t;
		case 'thisWeek':
			return t - v <= weekday - 1 && v - t <= 7 - weekday;
		case 'lastWeek':
			return t - v >= weekday && t - v < weekday + 7;
		case 'nextWeek':
			return v - t > 7 - weekday && v - t < 15 - weekday;
		case 'thisMonth':
			return monthOf(v) === current;
		case 'lastMonth':
			return monthOf(v) === current - 1;
		case 'nextMonth':
			return monthOf(v) === current + 1;
	}
}
