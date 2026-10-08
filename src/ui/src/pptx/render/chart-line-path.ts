/**
 * chart-line-path.ts: bezier smoothing for line-series paths.
 *
 * OOXML line/scatter series may set `c:smooth`, which PowerPoint renders as a
 * smoothed curve through the data points rather than straight segments. This
 * helper converts a point list into an SVG path `d` string using a Catmull-Rom
 * spline expressed as cubic beziers, so every binding can render the smoothed
 * line identically.
 *
 * @module chart-line-path
 */

import type { LinePoint } from './chart-view-model';

/**
 * Build a smoothed cubic-bezier path `d` through `points` (Catmull-Rom spline).
 * Returns a straight `M…L…` fallback for fewer than three points.
 */
export function smoothLinePath(points: ReadonlyArray<LinePoint>): string {
	if (points.length === 0) {
		return '';
	}
	if (points.length === 1) {
		return `M${points[0].x.toFixed(2)},${points[0].y.toFixed(2)}`;
	}
	if (points.length === 2) {
		return `M${points[0].x.toFixed(2)},${points[0].y.toFixed(2)} L${points[1].x.toFixed(2)},${points[1].y.toFixed(2)}`;
	}

	let d = `M${points[0].x.toFixed(2)},${points[0].y.toFixed(2)}`;
	for (let i = 0; i < points.length - 1; i++) {
		const [cp1, cp2, end] = smoothSegment(points, i);
		d += ` C${cp1.x.toFixed(2)},${cp1.y.toFixed(2)} ${cp2.x.toFixed(2)},${cp2.y.toFixed(2)} ${end.x.toFixed(2)},${end.y.toFixed(2)}`;
	}
	return d;
}

/** Bezier control points and end point of the Catmull-Rom segment from `points[i]`. */
function smoothSegment(
	points: ReadonlyArray<LinePoint>,
	i: number,
): [LinePoint, LinePoint, LinePoint] {
	const p0 = points[i - 1] ?? points[i];
	const p1 = points[i];
	const p2 = points[i + 1];
	const p3 = points[i + 2] ?? p2;
	return [
		{ x: p1.x + (p2.x - p0.x) / 6, y: p1.y + (p2.y - p0.y) / 6 },
		{ x: p2.x - (p3.x - p1.x) / 6, y: p2.y - (p3.y - p1.y) / 6 },
		p2,
	];
}

/**
 * Flatten the curve {@link smoothLinePath} draws into a polyline of
 * `stepsPerSegment` points per segment, so it can be clipped geometrically
 * (`chart-plot-clip.ts`). Fewer than three points are returned as they are,
 * since `smoothLinePath` draws those straight.
 */
export function sampleSmoothLine(
	points: ReadonlyArray<LinePoint>,
	stepsPerSegment = 16,
): LinePoint[] {
	if (points.length < 3) {
		return [...points];
	}
	const out: LinePoint[] = [points[0]];
	for (let i = 0; i < points.length - 1; i++) {
		const start = points[i],
			[cp1, cp2, end] = smoothSegment(points, i);
		for (let step = 1; step <= stepsPerSegment; step++) {
			const t = step / stepsPerSegment,
				u = 1 - t,
				a = u * u * u,
				b = 3 * u * u * t,
				c = 3 * u * t * t,
				d = t * t * t;
			out.push(
				step === stepsPerSegment
					? end
					: {
							x: a * start.x + b * cp1.x + c * cp2.x + d * end.x,
							y: a * start.y + b * cp1.y + c * cp2.y + d * end.y,
						},
			);
		}
	}
	return out;
}

/** An SVG path `d` drawing each run as its own straight-segment subpath. */
export function polylineRunsPath(runs: ReadonlyArray<ReadonlyArray<LinePoint>>): string {
	return runs
		.map((run) =>
			run.map((p, i) => `${i === 0 ? 'M' : 'L'}${p.x.toFixed(2)},${p.y.toFixed(2)}`).join(' '),
		)
		.join(' ');
}
