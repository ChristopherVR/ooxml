// Argument handling shared by the FORECAST.ETS functions: the timeline rules, aggregation of
// duplicate timestamps and data completion of missing ones.
import type { CallContext } from '../context';
import { ERR, fail, isError, type Value } from '../values';
import { timelineGrid } from './ets-grid';

export interface EtsSeries {
	/** Values on an evenly spaced grid. */
	y: number[];
	/** Position of a timeline value on the grid (0 is the first point); fractional between points. */
	position: (time: number) => number;
	/** The step in timeline units (the first interval in days for a calendar-month grid). */
	step: number;
}

const MAX_POINTS = 1_000_000;

/** The numbers of a range; an empty cell is `null` when `blanks` allows it. */
function numbersOf(ctx: CallContext, value: Value, blanks: boolean): (number | null)[] {
	return ctx
		.toMatrix(value)
		.flat()
		.map((cell) => {
			if (isError(cell)) fail(cell);
			if (cell === null && blanks) return null;
			if (typeof cell !== 'number') return fail(ERR.VALUE);
			return cell;
		});
}

const median = (xs: number[]): number => {
	const s = [...xs].sort((a, b) => a - b);
	const mid = s.length >> 1;
	return s.length % 2 ? (s[mid] as number) : ((s[mid - 1] as number) + (s[mid] as number)) / 2;
};

/** Aggregation codes 1..7: AVERAGE, COUNT, COUNTA, MAX, MIN, MEDIAN, SUM. */
function aggregate(code: number, xs: number[]): number {
	switch (code) {
		case 1:
			return xs.reduce((a, b) => a + b, 0) / xs.length;
		case 2:
		case 3:
			return xs.length;
		case 4:
			return Math.max(...xs);
		case 5:
			return Math.min(...xs);
		case 6:
			return median(xs);
		default:
			return xs.reduce((a, b) => a + b, 0);
	}
}

/** Validates `values` and `timeline` and returns the evenly spaced series (Excel's rules). */
export function etsSeries(
	ctx: CallContext,
	values: Value,
	timeline: Value,
	completion: number,
	aggregation: number,
): EtsSeries {
	if (!Number.isInteger(completion) || completion < 0 || completion > 1) fail(ERR.NUM);
	if (!Number.isInteger(aggregation) || aggregation < 1 || aggregation > 7) fail(ERR.NUM);
	const ys = numbersOf(ctx, values, true);
	const ts = numbersOf(ctx, timeline, false) as number[];
	if (ys.length !== ts.length) fail(ERR.NA);
	// An empty value is a missing point: its timestamp stays on the grid and is completed.
	const groups = new Map<number, number[]>();
	const stamps = new Set<number>(ts);
	ts.forEach((t, i) => {
		const v = ys[i];
		if (v === null || v === undefined) return;
		const group = groups.get(t) ?? [];
		group.push(v);
		groups.set(t, group);
	});
	const times = [...stamps].sort((a, b) => a - b);
	if (times.length < 2) fail(ERR.NUM);
	const grid = timelineGrid(times);
	const length = (grid.index[grid.index.length - 1] as number) + 1;
	if (length > MAX_POINTS) fail(ERR.NUM);
	const y: (number | undefined)[] = new Array<number | undefined>(length).fill(undefined);
	times.forEach((t, i) => {
		const group = groups.get(t);
		if (group) y[grid.index[i] as number] = aggregate(aggregation, group);
	});
	if (y[0] === undefined || y[length - 1] === undefined) {
		// Leading or trailing gaps cannot be interpolated; they take the nearest value (or zero).
		const first = y.findIndex((v) => v !== undefined);
		if (first < 0) fail(ERR.NUM);
		let last = length - 1;
		while (last > first && y[last] === undefined) last--;
		for (let i = 0; i < first; i++) y[i] = completion === 1 ? y[first] : 0;
		for (let i = last + 1; i < length; i++) y[i] = completion === 1 ? y[last] : 0;
	}
	// Missing grid points: 1 interpolates linearly between neighbours, 0 uses zero.
	let prev = 0;
	for (let i = 1; i < length; i++) {
		if (y[i] !== undefined) {
			if (i - prev > 1)
				for (let j = prev + 1; j < i; j++)
					y[j] =
						completion === 1
							? (y[prev] as number) +
								(((y[i] as number) - (y[prev] as number)) * (j - prev)) / (i - prev)
							: 0;
			prev = i;
		}
	}
	return { y: y as number[], position: grid.position, step: grid.stepStat };
}
