// What a chart's declared point counts cost. The model is sparse: a cache holds the points the file
// really has (`points`) and the count it declares (`pointCount`), so reading and saving scale with
// the bytes present and a declared count survives exactly as written. Only the dense views hosts
// draw and edit, one slot per position, expand blanks, and those are paid for from a budget a whole
// package shares. A view the budget cannot cover loses its trailing blanks, then its far points.
import type { XmlElement } from '../xml/index';
import type { ChartParseContext } from './parse-util';

/** Blank slots the chart views of one package may expand to; points in the file never count. */
export const MAX_CHART_BLANK_POINTS = 4_194_304;

/** Blank slots one view may hold: a whole worksheet column. */
export const MAX_CACHE_BLANK_POINTS = 1_048_576;

/** `ChartParseIssue.code` of a part with a view shorter than its declared points. */
export const CHART_POINTS_LIMITED = 'CHART_POINTS_LIMITED';

/** The blank slots a package has left, and how many views did not fit. */
export interface ChartPointBudget {
	blanks: number;
	cuts: number;
}

export const chartPointBudget = (blanks = MAX_CHART_BLANK_POINTS): ChartPointBudget => ({
	blanks,
	cuts: 0,
});

interface ChartPointSlot {
	index?: number | undefined;
}

/** A sparse cache, with the length of its view when the budget cut it. */
export interface ChartPointCache<P extends ChartPointSlot = ChartPointSlot> {
	points: readonly P[];
	pointCount?: number | undefined;
	viewLength?: number | undefined;
}

/** A point's slot: its own index, else the one after the highest slot so far. */
const slotOf = (point: ChartPointSlot, next: number) =>
	point.index !== undefined && Number.isInteger(point.index) && point.index >= 0
		? point.index
		: next;

/** One past the highest slot `points` occupy. */
export function chartPointsEnd(points: readonly ChartPointSlot[]): number {
	let end = 0;
	for (const point of points) end = Math.max(end, slotOf(point, end) + 1);
	return end;
}

const fullLength = (end: number, declared: number | undefined) =>
	declared !== undefined && Number.isFinite(declared) ? Math.max(end, declared) : end;

/** The length of a view no budget has cut: every slot, up to one column of blanks. */
export const chartPointLength = (present: number, end: number, declared: number | undefined) =>
	Math.min(fullLength(end, declared), present + MAX_CACHE_BLANK_POINTS);

const CACHE_ELEMENTS = ['numCache', 'strCache', 'numLit', 'strLit', 'lvl'];

/**
 * The blank slots each view of a parsed part may hold. Hosts pad a chart's series to its longest
 * one, so what a part can cost is its longest view times its caches, counted here by local name.
 */
export const chartViewAllowance = (root: XmlElement): number =>
	Math.floor(
		MAX_CHART_BLANK_POINTS /
			Math.max(
				1,
				CACHE_ELEMENTS.reduce(
					(total, name) => total + root.getElementsByTagNameNS('*', name).length,
					0,
				),
			),
	);

/** Pays for `blanks` slots of one view or grid, when the budget can cover them. */
export function spendChartPointBlanks(
	budget: ChartPointBudget,
	blanks: number,
	allowance = MAX_CACHE_BLANK_POINTS,
): boolean {
	if (blanks > Math.min(budget.blanks, allowance, MAX_CACHE_BLANK_POINTS)) return false;
	budget.blanks -= Math.max(0, blanks);
	return true;
}

/** The length of a view paid for from `budget`, counting a view it cannot cover in `cuts`. */
export function limitChartPointLength(
	budget: ChartPointBudget,
	present: number,
	end: number,
	declared: number | undefined,
	allowance?: number,
): number {
	const full = fullLength(end, declared);
	if (spendChartPointBlanks(budget, full - present, allowance)) return full;
	budget.cuts += 1;
	return spendChartPointBlanks(budget, end - present, allowance) ? end : present;
}

/**
 * The `viewLength` of a cache being parsed, absent when its view shows every declared slot. The
 * first cut in a part adds {@link CHART_POINTS_LIMITED}; the points and count stay as written.
 */
export function parsedChartViewLength(
	context: ChartParseContext,
	points: readonly ChartPointSlot[],
	declared: number | undefined,
): number | undefined {
	const budget = (context.pointBudget ??= chartPointBudget());
	const cuts = budget.cuts;
	const end = chartPointsEnd(points);
	const length = limitChartPointLength(budget, points.length, end, declared, context.viewAllowance);
	if (budget.cuts === cuts) return undefined;
	if (!context.pointsLimited) {
		context.pointsLimited = true;
		context.issues.push({
			code: CHART_POINTS_LIMITED,
			message:
				'The chart declares more blank points than are shown; its points and counts are kept.',
		});
	}
	return length;
}

/** The dense length of `cache`; `declared` replaces its own count (multi-level categories). */
export const chartPointsLength = (
	cache: ChartPointCache,
	declared: number | undefined = cache.pointCount,
): number =>
	cache.viewLength ?? chartPointLength(cache.points.length, chartPointsEnd(cache.points), declared);

/** The points of `cache` with their slots, in file order. */
export function slottedPoints<P extends ChartPointSlot>(
	cache: Pick<ChartPointCache<P>, 'points'>,
): [number, P][] {
	let next = 0;
	return cache.points.map((point) => {
		const slot = slotOf(point, next);
		next = Math.max(next, slot + 1);
		return [slot, point];
	});
}

/** The slots that hold a point, ascending: what a per-point edit has to visit. */
export const chartPointSlots = (cache: ChartPointCache): number[] =>
	[...new Set(slottedPoints(cache).map(([slot]) => slot))].sort((a, b) => a - b);

/** The points past the end of a cut view: an edit of the view keeps them. */
export const chartPointsPastView = <P extends ChartPointSlot>(cache: ChartPointCache<P>): P[] =>
	cache.viewLength === undefined
		? []
		: slottedPoints(cache).flatMap(([slot, point]) => (slot < cache.viewLength! ? [] : [point]));

/** `cache` holding `entries` in slot order, each point carrying its slot. */
function withSlots<P extends ChartPointSlot, C extends ChartPointCache<P>>(
	cache: C,
	entries: readonly (readonly [number, P])[],
	pointCount: number | undefined,
	viewLength: number | undefined,
): C {
	const next = { ...cache } as C & { points: P[]; pointCount?: number; viewLength?: number };
	next.points = [...entries]
		.sort((a, b) => a[0] - b[0])
		.map(([slot, point]) => (point.index === slot ? point : { ...point, index: slot }));
	if (pointCount === undefined) delete next.pointCount;
	else next.pointCount = pointCount;
	if (viewLength === undefined) delete next.viewLength;
	else next.viewLength = viewLength;
	return next;
}

const reaching = (length: number | undefined, slot: number) =>
	length === undefined ? undefined : Math.max(length, slot + 1);

/** `cache` with `point` in `slot`, or with that slot emptied. Every other point stays. */
export function chartCacheWithPoint<P extends ChartPointSlot, C extends ChartPointCache<P>>(
	cache: C,
	slot: number,
	point: P | undefined,
): C {
	const kept = slottedPoints<P>(cache).filter(([at]) => at !== slot);
	return withSlots<P, C>(
		cache,
		point ? [...kept, [slot, point]] : kept,
		point ? reaching(cache.pointCount, slot) : cache.pointCount,
		reaching(cache.viewLength, slot),
	);
}

/**
 * `cache` with one more row at `slot`, which holds `point` unless the file already has a point
 * there. A declared range that reaches the slot keeps its count.
 */
export function chartCacheWithRow<P extends ChartPointSlot, C extends ChartPointCache<P>>(
	cache: C,
	slot: number,
	point: P | undefined,
): C {
	const entries = slottedPoints<P>(cache);
	return withSlots<P, C>(
		cache,
		point && !entries.some(([at]) => at === slot) ? [...entries, [slot, point]] : entries,
		reaching(cache.pointCount, slot),
		reaching(cache.viewLength, slot),
	);
}

/**
 * `cache` without the row at `slot`: later points move up one. A cut view keeps its declared count,
 * because the range the count describes is still there; a complete cache gets one slot shorter.
 */
export function chartCacheWithoutRow<P extends ChartPointSlot, C extends ChartPointCache<P>>(
	cache: C,
	slot: number,
): C {
	const { pointCount, viewLength } = cache;
	const entries = slottedPoints<P>(cache).flatMap(([at, point]): [number, P][] =>
		at === slot ? [] : [[at > slot ? at - 1 : at, point]],
	);
	const shorter = (length: number | undefined) =>
		length !== undefined && slot < length ? length - 1 : length;
	return withSlots<P, C>(
		cache,
		entries,
		viewLength === undefined ? shorter(pointCount) : pointCount,
		shorter(viewLength),
	);
}

/**
 * The count and points of a cache rewritten from a whole view of `length` slots. A cut view keeps
 * the declared count and the points past its end: a shorter view lost rows, so those points move up
 * with it, and a longer one takes only the slots the file leaves free.
 */
export function chartCacheOverView<P extends ChartPointSlot>(
	written: P[],
	length: number,
	previous: ChartPointCache<P> | undefined,
): { pointCount?: number; points: P[]; viewLength?: number } {
	const view = previous?.viewLength;
	if (previous === undefined || view === undefined) return { pointCount: length, points: written };
	const removed = Math.max(0, view - length);
	const past = slottedPoints(previous).flatMap(([slot, point]): [number, P][] =>
		slot < view ? [] : [[slot - removed, point]],
	);
	const taken = new Set(past.map(([slot]) => slot));
	const fresh = slottedPoints<P>({ points: written }).filter(([slot]) => !taken.has(slot));
	const viewLength = fresh.reduce((end, [slot]) => Math.max(end, slot + 1), Math.min(view, length));
	const { points } = withSlots<P, ChartPointCache<P>>(
		{ points: [] },
		[...fresh, ...past],
		undefined,
		undefined,
	);
	return {
		...(previous.pointCount === undefined
			? {}
			: { pointCount: Math.max(previous.pointCount, viewLength) }),
		points: points as P[],
		viewLength,
	};
}

/** `cache` laid out by slot; slots without a point hold `blank`, later points win a slot. */
export function chartDensePoints<P extends ChartPointSlot, T>(
	cache: ChartPointCache<P>,
	read: (point: P) => T,
	blank: T,
): T[] {
	const dense = new Array<T>(chartPointsLength(cache)).fill(blank);
	for (const [slot, point] of slottedPoints(cache))
		if (slot < dense.length) dense[slot] = read(point);
	return dense;
}
