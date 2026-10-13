import type { VisioRoutePoint } from './connector-route';
import { flattenNurbs } from './nurbs-flatten';

const mix = (a: VisioRoutePoint, b: VisioRoutePoint, t: number): VisioRoutePoint => ({
	x: a.x + (b.x - a.x) * t,
	y: a.y + (b.y - a.y) * t,
});

/** Native single-corner Dynamic connector: five controls and one internal cubic knot. */
export function cornerCurve(path: readonly VisioRoutePoint[]): VisioRoutePoint[] {
	const [begin, corner, end] = path as [VisioRoutePoint, VisioRoutePoint, VisioRoutePoint];
	return [begin, mix(begin, corner, 0.25), corner, mix(corner, end, 0.75), end];
}

/** Distances and compact knot values recorded from native Visio's corner curves. */
export function cornerCurveKnots(points: readonly VisioRoutePoint[]): [number, number] {
	const distance = (a: number, b: number) =>
		Math.hypot(points[b]!.x - points[a]!.x, points[b]!.y - points[a]!.y) / 3;
	const internal = distance(0, 3);
	return [internal, internal + distance(1, 4)];
}

/** Insert the internal knot twice to obtain exact, ordinary cubic Bezier spans. */
export function connectorCurveSpans(points: readonly VisioRoutePoint[]): VisioRoutePoint[][] {
	if (points.length === 4) return [[...points]];
	if (points.length !== 5) return [];
	const [internal, last] = cornerCurveKnots(points);
	const ratio = internal / last;
	const left = mix(points[1]!, points[2]!, ratio);
	const right = mix(points[2]!, points[3]!, ratio);
	const join = mix(left, right, ratio);
	return [
		[points[0]!, points[1]!, left, join],
		[join, right, points[3]!, points[4]!],
	];
}

/** The point halfway along a nonempty polyline, used for connector text placement. */
export function connectorPolylineMiddle(points: readonly VisioRoutePoint[]): VisioRoutePoint {
	const lengths = points.slice(1).map((p, i) => Math.hypot(p.x - points[i]!.x, p.y - points[i]!.y));
	let remaining = lengths.reduce((sum, length) => sum + length, 0) / 2;
	for (const [i, length] of lengths.entries()) {
		if (remaining <= length && length > 0)
			return mix(points[i]!, points[i + 1]!, remaining / length);
		remaining -= length;
	}
	return points.at(-1)!;
}

/** Sample the native corner geometry for text placement and conservative collision checks. */
export function flattenCornerCurve(
	points: readonly VisioRoutePoint[],
): VisioRoutePoint[] | undefined {
	const [internal, last] = cornerCurveKnots(points);
	const flattened = flattenNurbs(
		points.map((p) => ({ ...p, weight: 1 })),
		[0, 0, 0, 0, internal, last, last, last, last],
		3,
		() => {},
		() => {},
	);
	return flattened ? [points[0]!, ...flattened] : undefined;
}

/** Estimate the corner curve's arc midpoint using the shared bounded NURBS flattener. */
export function connectorCurveMiddle(points: readonly VisioRoutePoint[]): VisioRoutePoint {
	const spans = connectorCurveSpans(points);
	if (points.length === 5) {
		const flattened = flattenCornerCurve(points);
		if (flattened) return connectorPolylineMiddle(flattened);
		return mix(points[0]!, points.at(-1)!, 0.5);
	}
	const row = [...spans[0]!];
	for (let count = row.length - 1; count > 0; count--)
		for (let i = 0; i < count; i++) row[i] = mix(row[i]!, row[i + 1]!, 0.5);
	return row[0]!;
}
