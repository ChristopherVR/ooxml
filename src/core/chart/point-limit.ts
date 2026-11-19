// Limits on the dense arrays built from chart point caches. `c:ptCount`, `cx:lvl/@ptCount` and
// `@idx` are claims a few bytes long, so memory follows the points a file really holds: a cache
// expands to its present points plus a bounded number of blank slots. A part shares one allowance
// among its caches, because series are padded to the longest one when a chart is laid out.
import type { XmlElement } from '../xml/index';
import type { ChartParseContext } from './parse-util';

/**
 * Blank slots one chart part may expand to, split evenly among its caches. Holes between points
 * and padding up to the declared count are each held to a cache's share; points present in the
 * file never count.
 */
export const MAX_CHART_BLANK_POINTS = 262_144;

/** `ChartParseIssue.code` of a part whose declared points were cut to the limit. */
export const CHART_POINTS_LIMITED = 'CHART_POINTS_LIMITED';

/** The share of blank slots each cache of a part may use, and whether any cache went over it. */
export interface ChartPointLimit {
	blanks: number;
	limited: boolean;
}

export const chartPointLimit = (caches = 1): ChartPointLimit => ({
	blanks: Math.floor(MAX_CHART_BLANK_POINTS / Math.max(1, caches)),
	limited: false,
});

const CACHE_ELEMENTS = ['numCache', 'strCache', 'numLit', 'strLit', 'lvl', 'dlblRangeCache'];
const CACHE_TAG = new RegExp(`<(?:[\\w.-]+:)?(?:${CACHE_ELEMENTS.join('|')})[\\s/>]`, 'gu');

/** The limit of a parsed chart part, shared among every point cache in it. */
export const chartPointLimitOfPart = (root: XmlElement): ChartPointLimit =>
	chartPointLimit(
		CACHE_ELEMENTS.reduce(
			(total, name) => total + root.getElementsByTagNameNS('*', name).length,
			0,
		),
	);

/** {@link chartPointLimitOfPart} for a part still in text form. */
export const chartPointLimitOfXml = (xml: string): ChartPointLimit =>
	chartPointLimit(xml.match(CACHE_TAG)?.length ?? 0);

interface ChartPointSlot {
	index?: number | undefined;
}

/** A point's slot: its own index, else the one after the highest slot so far. */
const slotOf = (point: ChartPointSlot, next: number) =>
	point.index !== undefined && Number.isInteger(point.index) && point.index >= 0
		? point.index
		: next;

function pointsEnd(points: readonly ChartPointSlot[]): number {
	let end = 0;
	for (const point of points) end = Math.max(end, slotOf(point, end) + 1);
	return end;
}

/**
 * The dense length of a cache: `end` (one past its highest point) or `declared` when longer, with
 * at most `blanks` holes below the points and `blanks` padding slots after them.
 */
export function chartPointLength(
	present: number,
	end: number,
	declared: number | undefined,
	blanks = MAX_CHART_BLANK_POINTS,
): number {
	const filled = Math.min(end, present + blanks);
	const wanted = Number.isFinite(declared) ? Math.max(filled, declared!) : filled;
	return Math.min(wanted, filled + blanks);
}

/** {@link chartPointLength} under a part's limit, recording when the cache does not fit. */
export function limitChartPointLength(
	limit: ChartPointLimit,
	present: number,
	end: number,
	declared: number | undefined,
): number {
	const length = chartPointLength(present, end, declared, limit.blanks);
	if (length < end || (Number.isFinite(declared) && length < declared!)) limit.limited = true;
	return length;
}

/** The dense length of `points` with `declared` slots. */
export const chartPointsLength = (
	points: readonly ChartPointSlot[],
	declared: number | undefined,
) => chartPointLength(points.length, pointsEnd(points), declared);

/** `points` laid out by slot; slots without a point hold `blank`, later points win a slot. */
export function chartDensePoints<P extends ChartPointSlot, T>(
	points: readonly P[],
	declared: number | undefined,
	read: (point: P) => T,
	blank: T,
): T[] {
	const dense = new Array<T>(chartPointsLength(points, declared)).fill(blank);
	let next = 0;
	for (const point of points) {
		const slot = slotOf(point, next);
		next = Math.max(next, slot + 1);
		if (slot < dense.length) dense[slot] = read(point);
	}
	return dense;
}

/**
 * Bounds one parsed cache by its part's limit. A cache within it keeps its points and count as
 * written; the first cut adds {@link CHART_POINTS_LIMITED} to the issues, so hosts keep the part's
 * source bytes instead of rewriting it from the shortened model.
 */
export function limitParsedChartPoints<P extends ChartPointSlot>(
	context: ChartParseContext,
	points: P[],
	declared: number | undefined,
): { points: P[]; pointCount: number | undefined } {
	const limit = (context.pointLimit ??= chartPointLimit());
	const wasLimited = limit.limited;
	const end = pointsEnd(points);
	const length = limitChartPointLength(limit, points.length, end, declared);
	if (limit.limited && !wasLimited)
		context.issues.push({
			code: CHART_POINTS_LIMITED,
			message: `The chart declares more blank points than the ${MAX_CHART_BLANK_POINTS} read per part; the rest are left out.`,
		});
	let next = 0;
	const kept =
		end <= length
			? points
			: points.filter((point) => {
					const slot = slotOf(point, next);
					next = Math.max(next, slot + 1);
					return slot < length;
				});
	return {
		points: kept,
		pointCount: declared === undefined ? undefined : Math.min(declared, length),
	};
}
