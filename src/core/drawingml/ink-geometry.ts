export interface InkCoordinate {
	x: number;
	y: number;
}

export const MAX_INK_SAMPLES = 4096;

export function inkDistance(a: InkCoordinate, b: InkCoordinate): number {
	return Math.hypot(a.x - b.x, a.y - b.y);
}

export function inkBoundsOf(points: readonly InkCoordinate[]) {
	let x = Infinity;
	let y = Infinity;
	let right = -Infinity;
	let bottom = -Infinity;
	for (const point of points) {
		x = Math.min(x, point.x);
		y = Math.min(y, point.y);
		right = Math.max(right, point.x);
		bottom = Math.max(bottom, point.y);
	}
	return { x, y, width: right - x, height: bottom - y };
}

/** Project onto an infinite line. The returned offset is in units of its length. */
export function projectInkPoint<T extends InkCoordinate>(
	point: T,
	a: InkCoordinate,
	b: InkCoordinate,
) {
	const dx = b.x - a.x;
	const dy = b.y - a.y;
	const lengthSquared = dx * dx + dy * dy;
	const offset =
		lengthSquared === 0 ? 0 : ((point.x - a.x) * dx + (point.y - a.y) * dy) / lengthSquared;
	return { offset, point: { ...point, x: a.x + offset * dx, y: a.y + offset * dy } };
}

function segmentDistance(point: InkCoordinate, a: InkCoordinate, b: InkCoordinate) {
	const projection = projectInkPoint(point, a, b);
	return inkDistance(
		point,
		projection.offset < 0 ? a : projection.offset > 1 ? b : projection.point,
	);
}

/** Iterative simplification keeps stack depth independent of sample count. */
export function simplifyInk(points: readonly InkCoordinate[], tolerance: number): InkCoordinate[] {
	const keep = new Set([0, points.length - 1]);
	const segments: [number, number][] = [[0, points.length - 1]];
	for (let segment = segments.pop(); segment; segment = segments.pop()) {
		const [first, last] = segment;
		const a = points[first];
		const b = points[last];
		if (!a || !b) {
			continue;
		}
		let furthest = first;
		let distance = tolerance;
		for (let index = first + 1; index < last; index++) {
			const point = points[index];
			if (!point) {
				continue;
			}
			const error = segmentDistance(point, a, b);
			if (error > distance) {
				distance = error;
				furthest = index;
			}
		}
		if (furthest > first) {
			keep.add(furthest);
			segments.push([first, furthest], [furthest, last]);
		}
	}
	return points.filter((_, index) => keep.has(index));
}
