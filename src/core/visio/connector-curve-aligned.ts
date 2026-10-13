import type { VisioWalkEnd } from './connector-route-walk';
import type { VisioRoutePoint } from './connector-route';

/** Native curved walking glue uses the midpoint of the overlapping facing sides. */
export function alignedVisioCurve(
	begin: VisioWalkEnd,
	end: VisioWalkEnd,
): VisioRoutePoint[] | undefined {
	const a = begin.shape;
	const b = end.shape;
	if (!a || !b) return undefined;
	// Bounds are not a valid outline for rotated targets.
	for (const shape of [a, b])
		if (
			shape.corners.some(
				(p) =>
					(Math.abs(p.x - shape.bounds.minX) > 1e-9 && Math.abs(p.x - shape.bounds.maxX) > 1e-9) ||
					(Math.abs(p.y - shape.bounds.minY) > 1e-9 && Math.abs(p.y - shape.bounds.maxY) > 1e-9),
			)
		)
			return undefined;
	for (const axis of ['x', 'y'] as const) {
		const horizontal = axis === 'x';
		const low = horizontal ? 'minY' : 'minX';
		const high = horizontal ? 'maxY' : 'maxX';
		const overlapLow = Math.max(a.bounds[low], b.bounds[low]);
		const overlapHigh = Math.min(a.bounds[high], b.bounds[high]);
		if (overlapHigh - overlapLow <= 1e-9) continue;
		const forward = b.center[axis] > a.center[axis];
		const from = horizontal
			? a.bounds[forward ? 'maxX' : 'minX']
			: a.bounds[forward ? 'maxY' : 'minY'];
		const to = horizontal
			? b.bounds[forward ? 'minX' : 'maxX']
			: b.bounds[forward ? 'minY' : 'maxY'];
		if ((to - from) * (forward ? 1 : -1) <= 1e-9) continue;
		const middle = (overlapLow + overlapHigh) / 2;
		return [0, 0.25, 0.75, 1].map((t) => {
			const along = from + (to - from) * t;
			return horizontal ? { x: along, y: middle } : { x: middle, y: along };
		});
	}
	return undefined;
}
