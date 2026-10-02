import type { Cell, CellValue, Workbook } from '../model.js';
import { styleAt } from '../styles.js';
import { dateToSerial, isDateFormat, serialToDate } from './deps.js';

/** What a fill puts at lane position `k`: a copy of the source cell or a generated value. */
export type SeriesStep = { kind: 'copy' } | { kind: 'value'; value: CellValue };

export type SeriesGenerator = (k: number) => SeriesStep;

const COPY: SeriesGenerator = () => ({ kind: 'copy' });

export const mod = (n: number, m: number): number => ((n % m) + m) % m;

const LISTS: string[][] = [
	['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'],
	['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'],
	['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],
	[
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
	],
];

function listMatch(text: string): { list: string[]; index: number } | undefined {
	const lower = text.toLowerCase();
	for (const list of LISTS) {
		const index = list.findIndex((item) => item.toLowerCase() === lower);
		if (index >= 0) return { list, index };
	}
	return undefined;
}

/** Re-applies the letter case of `sample` (`MON`, `mon`, `Mon`) to a list item. */
function matchCase(item: string, sample: string): string {
	if (sample === sample.toUpperCase()) return item.toUpperCase();
	if (sample === sample.toLowerCase()) return item.toLowerCase();
	return item;
}

/** Least-squares line through `(i, values[i])`, as Excel's AutoFill trend uses. */
function linearTrend(values: number[]): (k: number) => number {
	const n = values.length;
	const meanX = (n - 1) / 2;
	const meanY = values.reduce((a, b) => a + b, 0) / n;
	let num = 0;
	let den = 0;
	values.forEach((y, x) => {
		num += (x - meanX) * (y - meanY);
		den += (x - meanX) ** 2;
	});
	const slope = den ? num / den : 0;
	// Round away floating noise (0.1 + 0.2) the way Excel's 15 significant digits do.
	return (k) => Number((meanY + slope * (k - meanX)).toPrecision(15));
}

const TEXT_NUMBER = /^(.*?)(\d+)(\D*)$/;
const QUARTER = /^(q|qtr|quarter)(\s?)([1-4])$/i;

function textSeries(texts: string[]): SeriesGenerator | undefined {
	const lists = texts.map(listMatch);
	const first = lists[0];
	if (first && lists.every((m) => m && m.list === first.list)) {
		const second = lists[1];
		const step = second ? mod(second.index - first.index, first.list.length) || 1 : 1;
		const sample = texts[0] ?? '';
		return (k) => ({
			kind: 'value',
			value: matchCase(first.list[mod(first.index + step * k, first.list.length)] ?? '', sample),
		});
	}
	const quarters = texts.map((t) => QUARTER.exec(t));
	const q0 = quarters[0];
	if (q0 && quarters.every((q) => q && q[1] === q0[1] && q[2] === q0[2])) {
		const start = Number(q0[3]) - 1;
		const q1 = quarters[1];
		const step = q1 ? mod(Number(q1[3]) - 1 - start, 4) || 1 : 1;
		return (k) => ({ kind: 'value', value: `${q0[1]}${q0[2]}${mod(start + step * k, 4) + 1}` });
	}
	const parts = texts.map((t) => TEXT_NUMBER.exec(t));
	const p0 = parts[0];
	if (!p0 || !parts.every((p) => p && p[1] === p0[1] && p[3] === p0[3])) return undefined;
	const numbers = parts.map((p) => Number(p?.[2]));
	const width = (p0[2] ?? '').startsWith('0') ? (p0[2] ?? '').length : 0;
	const trend = numbers.length === 1 ? (k: number) => (numbers[0] ?? 0) + k : linearTrend(numbers);
	return (k) => {
		const n = Math.abs(Math.round(trend(k)));
		return { kind: 'value', value: `${p0[1]}${String(n).padStart(width, '0')}${p0[3]}` };
	};
}

function addMonths(serial: number, months: number, date1904: boolean): number {
	const date = serialToDate(serial, date1904);
	const y = date.getUTCFullYear();
	const m = date.getUTCMonth() + months;
	const lastDay = new Date(Date.UTC(y, m + 1, 0)).getUTCDate();
	const day = Math.min(date.getUTCDate(), lastDay);
	const next = new Date(
		Date.UTC(y, m, day, date.getUTCHours(), date.getUTCMinutes(), date.getUTCSeconds()),
	);
	return dateToSerial(next, date1904);
}

function dateSeries(serials: number[], date1904: boolean): SeriesGenerator {
	if (serials.length === 1) {
		const first = serials[0] ?? 0;
		return (k) => ({ kind: 'value', value: first + k });
	}
	const dates = serials.map((s) => serialToDate(s, date1904));
	const monthIndex = dates.map((d) => d.getUTCFullYear() * 12 + d.getUTCMonth());
	const sameDay = dates.every((d) => d.getUTCDate() === dates[0]?.getUTCDate());
	const stepMonths = (monthIndex[1] ?? 0) - (monthIndex[0] ?? 0);
	const evenMonths =
		stepMonths !== 0 && monthIndex.every((m, i) => m - (monthIndex[0] ?? 0) === stepMonths * i);
	const first = serials[0] ?? 0;
	if (sameDay && evenMonths)
		return (k) => ({ kind: 'value', value: addMonths(first, stepMonths * k, date1904) });
	const trend = linearTrend(serials);
	return (k) => ({ kind: 'value', value: trend(k) });
}

/**
 * Detects the series a lane of source cells forms (one column for a vertical fill, one row for
 * a horizontal one) and returns how to continue it. Formulas, blanks and mixed content copy.
 */
export function detectSeries(workbook: Workbook, lane: (Cell | undefined)[]): SeriesGenerator {
	if (!lane.length || lane.some((c) => !c || c.formula !== undefined || c.value === null))
		return COPY;
	const values = lane.map((c) => c?.value ?? null);
	if (values.every((v): v is number => typeof v === 'number')) {
		const isDate = isDateFormat(styleAt(workbook, lane[0]?.styleId).numFmt);
		if (isDate) return dateSeries(values, workbook.date1904);
		if (values.length === 1) return COPY;
		const trend = linearTrend(values);
		return (k) => ({ kind: 'value', value: trend(k) });
	}
	if (values.every((v): v is string => typeof v === 'string')) return textSeries(values) ?? COPY;
	return COPY;
}
