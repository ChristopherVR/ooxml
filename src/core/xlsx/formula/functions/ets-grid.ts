// The evenly spaced grid of a FORECAST.ETS timeline. Plain numbers step by the smallest gap; a date
// timeline on the same day of each month (or on month ends) with unequal gaps in days steps by
// calendar months, so monthly and yearly data work as in Excel.
import { serialToYmd, ymdToSerial } from '../date-serial';
import { ERR, fail } from '../values';

export interface TimelineGrid {
	/** Grid units between neighbouring points (days, or months on a calendar-month grid). */
	step: number;
	/** Grid position of each sorted time, relative to the first (0, ...). */
	index: number[];
	/** Position of any time on the grid; fractional between points. */
	position: (time: number) => number;
	/** The timeline value at grid position `i` (0 is the first time). */
	timeAt: (i: number) => number;
	/** The step reported by FORECAST.ETS.STAT (the first interval in days on a month grid). */
	stepStat: number;
}

const monthEnd = (d: { year: number; month: number; day: number }): boolean =>
	d.day === new Date(Date.UTC(d.year, d.month, 0)).getUTCDate();

/**
 * Month numbers of a date timeline that falls on the same day of each month (or on month ends)
 * with unequal gaps in days; `undefined` for any other timeline.
 */
function monthNumbers(times: readonly number[]): number[] | undefined {
	if (times.length < 2 || times.some((t) => !Number.isInteger(t) || t < 1)) return undefined;
	const gaps = new Set<number>();
	for (let i = 1; i < times.length; i++) gaps.add((times[i] as number) - (times[i - 1] as number));
	if (gaps.size < 2 || Math.min(...gaps) < 28) return undefined;
	const dates = times.map((t) => serialToYmd(t));
	const sameDay = dates.every((d) => d.day === (dates[0] as { day: number }).day);
	if (!sameDay && !dates.every(monthEnd)) return undefined;
	return dates.map((d) => d.year * 12 + d.month);
}

/** The grid of sorted, distinct `times`; fails with #NUM! when they are not evenly spaced. */
export function timelineGrid(times: readonly number[]): TimelineGrid {
	const months = monthNumbers(times);
	const axis = months ?? times;
	let step = Number.POSITIVE_INFINITY;
	for (let i = 1; i < axis.length; i++)
		step = Math.min(step, (axis[i] as number) - (axis[i - 1] as number));
	if (!(step > 0) || !Number.isFinite(step)) fail(ERR.NUM);
	const start = axis[0] as number;
	const index = axis.map((a) => {
		const k = (a - start) / step;
		if (Math.abs(k - Math.round(k)) > 1e-6) fail(ERR.NUM);
		return Math.round(k);
	});
	const first = times[0] as number;
	const firstDay = serialToYmd(first).day;
	const position = (time: number): number => {
		if (!months) return (time - start) / step;
		const d = serialToYmd(time);
		return (d.year * 12 + d.month - start) / step;
	};
	const timeAt = (i: number): number => {
		if (!months) return start + i * step;
		const total = start + i * step - 1;
		const year = Math.floor(total / 12);
		const month = total - year * 12 + 1;
		const day = monthEnd(serialToYmd(first))
			? new Date(Date.UTC(year, month, 0)).getUTCDate()
			: Math.min(firstDay, new Date(Date.UTC(year, month, 0)).getUTCDate());
		return ymdToSerial(year, month, day) ?? Number.NaN;
	};
	const stepStat = months ? (times[1] as number) - first : step;
	return { step, index, position, timeAt, stepStat };
}
