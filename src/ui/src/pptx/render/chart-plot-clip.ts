/**
 * chart-plot-clip.ts: geometric clipping of line and area series to the plot
 * area's vertical extent.
 *
 * An authored `c:min` / `c:max` can put a value (or a stacked running total)
 * outside the value axis. PowerPoint cuts the series at the plot edge. The
 * chart view model has no clip-path primitive, and adding one would need every
 * binding to render it, so the shared builders clip the geometry itself: a
 * line is split into the runs that stay inside the plot, and an area polygon
 * is cut by the top and bottom edges (Sutherland-Hodgman). Geometry already
 * inside the plot comes back unchanged, point for point.
 *
 * @module chart-plot-clip
 */
import { polylineRunsPath, sampleSmoothLine, smoothLinePath } from './chart-line-path';
import type { SvgPath, SvgPolyline } from './chart-svg-primitives';
import { linePointsToSvgString } from './chart-view-model-bars';
import type { LinePoint } from './chart-view-model-bars';
import type { PlotLayout } from './chart-view-model-types';

/** Vertical extent (pixel y) a series is clipped to. */
export interface PlotBand {
	top: number;
	bottom: number;
}

/** Floating-point slack so a point computed exactly on an edge counts as inside. */
const EPSILON = 1e-6;

/** The plot area's vertical extent. */
export function plotBand(layout: Pick<PlotLayout, 'plotTop' | 'plotBottom'>): PlotBand {
	return {
		top: Math.min(layout.plotTop, layout.plotBottom),
		bottom: Math.max(layout.plotTop, layout.plotBottom),
	};
}

/** Whether a point lies inside the band (edges included). */
export function insideBand(point: LinePoint, band: PlotBand): boolean {
	return point.y >= band.top - EPSILON && point.y <= band.bottom + EPSILON;
}

/**
 * Clip one segment to the band (Liang-Barsky on y). Returns the visible part,
 * reusing `a` / `b` themselves for an end that is not cut, or `undefined` when
 * nothing of it is visible.
 */
function clipSegment(
	a: LinePoint,
	b: LinePoint,
	band: PlotBand,
): [LinePoint, LinePoint] | undefined {
	const dy = b.y - a.y;
	if (Math.abs(dy) < EPSILON) {
		return insideBand(a, band) ? [a, b] : undefined;
	}
	const tTop = (band.top - a.y) / dy,
		tBottom = (band.bottom - a.y) / dy,
		low = Math.max(0, Math.min(tTop, tBottom)),
		high = Math.min(1, Math.max(tTop, tBottom));
	if (high - low < EPSILON) {
		return undefined;
	}
	const at = (t: number): LinePoint =>
		t <= 0 ? a : t >= 1 ? b : { x: a.x + (b.x - a.x) * t, y: a.y + dy * t };
	return [insideBand(a, band) ? a : at(low), insideBand(b, band) ? b : at(high)];
}

/**
 * Split an open polyline into the runs that lie inside the band. A line that
 * leaves the plot and comes back yields one run per visible stretch, each
 * ending where it crosses the edge. A line wholly inside comes back as one
 * run with its own points; runs of fewer than two points are dropped.
 */
export function clipPolylineToBand(
	points: ReadonlyArray<LinePoint>,
	band: PlotBand,
): LinePoint[][] {
	if (points.every((point) => insideBand(point, band))) {
		return points.length >= 2 ? [[...points]] : [];
	}
	const runs: LinePoint[][] = [];
	let run: LinePoint[] = [];
	const flush = (): void => {
		if (run.length >= 2) {
			runs.push(run);
		}
		run = [];
	};
	for (let i = 0; i + 1 < points.length; i++) {
		const clipped = clipSegment(points[i], points[i + 1], band);
		if (!clipped) {
			flush();
			continue;
		}
		const [start, end] = clipped;
		if (run[run.length - 1] !== start) {
			flush();
			run.push(start);
		}
		run.push(end);
		if (end !== points[i + 1]) {
			flush();
		}
	}
	flush();
	return runs;
}

/** Where the segment `a`-`b` crosses the horizontal line `y`. */
function crossAt(a: LinePoint, b: LinePoint, y: number): LinePoint {
	const t = (y - a.y) / (b.y - a.y);
	return { x: a.x + (b.x - a.x) * t, y };
}

/** One Sutherland-Hodgman pass against a horizontal edge. */
function clipPolygonEdge(
	polygon: ReadonlyArray<LinePoint>,
	keep: (point: LinePoint) => boolean,
	y: number,
): LinePoint[] {
	const out: LinePoint[] = [];
	for (let i = 0; i < polygon.length; i++) {
		const current = polygon[i],
			previous = polygon[(i + polygon.length - 1) % polygon.length];
		if (keep(current)) {
			if (!keep(previous)) {
				out.push(crossAt(previous, current, y));
			}
			out.push(current);
		} else if (keep(previous)) {
			out.push(crossAt(previous, current, y));
		}
	}
	return out;
}

/**
 * Clip a closed polygon (an area fill) to the band. A polygon wholly inside
 * comes back unchanged; one wholly outside comes back empty.
 */
export function clipPolygonToBand(polygon: ReadonlyArray<LinePoint>, band: PlotBand): LinePoint[] {
	if (polygon.every((point) => insideBand(point, band))) {
		return [...polygon];
	}
	const top = clipPolygonEdge(polygon, (point) => point.y >= band.top - EPSILON, band.top);
	return clipPolygonEdge(top, (point) => point.y <= band.bottom + EPSILON, band.bottom);
}

/**
 * Centre of the visible part of the vertical span `y1`..`y2` (a stacked band
 * at one category), or `undefined` when none of it is inside the band.
 */
export function visibleSpanCentre(y1: number, y2: number, band: PlotBand): number | undefined {
	const low = Math.max(Math.min(y1, y2), band.top),
		high = Math.min(Math.max(y1, y2), band.bottom);
	if (high < low - EPSILON) {
		return undefined;
	}
	return (low + high) / 2;
}

/**
 * The stroke of one line series, clipped to the band: a polyline per visible
 * run, or one path for a `c:smooth` curve (flattened only when it has to be
 * cut). A line wholly inside the band gives exactly the primitive it always
 * did; one wholly outside gives none.
 */
export function clippedSeriesLine(
	points: ReadonlyArray<LinePoint>,
	smooth: boolean,
	style: Omit<SvgPolyline, 'kind' | 'points'>,
	band: PlotBand,
): Array<SvgPath | SvgPolyline> {
	if (points.every((point) => insideBand(point, band))) {
		return [
			smooth
				? { kind: 'path', d: smoothLinePath(points), ...style }
				: { kind: 'polyline', points: linePointsToSvgString(points), ...style },
		];
	}
	if (smooth) {
		const runs = clipPolylineToBand(sampleSmoothLine(points), band);
		return runs.length > 0 ? [{ kind: 'path', d: polylineRunsPath(runs), ...style }] : [];
	}
	return clipPolylineToBand(points, band).map((run) => ({
		kind: 'polyline',
		points: linePointsToSvgString(run),
		...style,
	}));
}
