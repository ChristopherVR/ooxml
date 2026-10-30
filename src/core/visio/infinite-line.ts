import type { VisioMatrix } from './model';

/** The page rectangle [0, width] x [0, height] and the shape-local to page transform. */
export interface InfiniteLineFrame {
	/** Shape-local to page coordinates, composed through every ancestor group. */
	toPage: VisioMatrix;
	/** Page size in the same (pre-scale) units as the shape geometry. */
	width: number;
	height: number;
}

/**
 * Clips the infinite line through two shape-local points to the page rectangle.
 * Returns the visible endpoints in shape-local coordinates, or undefined when the points
 * coincide or the line misses the page. An affine map keeps the line parameter t, so the
 * clip is computed in page space and the same t is applied to the local points.
 */
export function clipInfiniteLine(
	x1: number,
	y1: number,
	x2: number,
	y2: number,
	frame: InfiniteLineFrame,
): [number, number, number, number] | undefined {
	const [a, b, c, d, e, f] = frame.toPage;
	const dx = x2 - x1,
		dy = y2 - y1;
	if (!(Math.hypot(dx, dy) > 0)) return undefined;
	const qx = a * x1 + c * y1 + e,
		qy = b * x1 + d * y1 + f;
	const ex = a * dx + c * dy,
		ey = b * dx + d * dy;
	if (!(Math.hypot(ex, ey) > 1e-12)) return undefined;
	// Liang-Barsky: each edge bounds t through p * t <= q.
	const edges: readonly (readonly [number, number])[] = [
		[-ex, qx],
		[ex, frame.width - qx],
		[-ey, qy],
		[ey, frame.height - qy],
	];
	let t0 = -Infinity,
		t1 = Infinity;
	for (const [p, q] of edges) {
		if (Math.abs(p) < 1e-12) {
			if (q < 0) return undefined;
			continue;
		}
		const r = q / p;
		if (p < 0) t0 = Math.max(t0, r);
		else t1 = Math.min(t1, r);
	}
	if (!(t0 <= t1)) return undefined;
	return [x1 + t0 * dx, y1 + t0 * dy, x1 + t1 * dx, y1 + t1 * dy];
}
