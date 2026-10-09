/** A point in any consistent planar coordinate system (inches). */
export interface VisioPathPoint {
	x: number;
	y: number;
}
/**
 * One drawn path segment ending at `x`, `y`. `cubic` carries two Bezier control points; `arc` is an
 * axis-aligned elliptical arc through the point (`a`, `b`) whose ellipse has `ratio` = rx / ry.
 */
export type VisioPathSegment =
	| { kind: 'line'; x: number; y: number }
	| { kind: 'cubic'; x: number; y: number; x1: number; y1: number; x2: number; y2: number }
	| { kind: 'arc'; x: number; y: number; a: number; b: number; ratio: number };

const finite = (point: VisioPathPoint) => Number.isFinite(point.x) && Number.isFinite(point.y);
const distance = (a: VisioPathPoint, b: VisioPathPoint) => Math.hypot(b.x - a.x, b.y - a.y);

/** Distance from `point` to the segment `a`-`b`. */
export function visioSegmentDistance(
	point: VisioPathPoint,
	a: VisioPathPoint,
	b: VisioPathPoint,
): number {
	const dx = b.x - a.x,
		dy = b.y - a.y,
		length = dx * dx + dy * dy;
	if (!length) return distance(point, a);
	const t = Math.max(0, Math.min(1, ((point.x - a.x) * dx + (point.y - a.y) * dy) / length));
	return Math.hypot(point.x - (a.x + t * dx), point.y - (a.y + t * dy));
}

/** Drop non-finite and consecutive duplicate samples. */
export function visioCleanSamples(points: readonly VisioPathPoint[]): VisioPathPoint[] {
	const result: VisioPathPoint[] = [];
	for (const point of points) {
		if (!finite(point)) continue;
		const last = result[result.length - 1];
		if (!last || distance(last, point) > 1e-9) result.push({ x: point.x, y: point.y });
	}
	return result;
}

/** Ramer-Douglas-Peucker simplification (iterative); keeps both ends. Returns kept indices. */
export function visioSimplifyIndices(
	points: readonly VisioPathPoint[],
	tolerance: number,
): number[] {
	if (points.length < 3) return points.map((_, index) => index);
	const keep = new Uint8Array(points.length);
	keep[0] = keep[points.length - 1] = 1;
	const stack: [number, number][] = [[0, points.length - 1]];
	while (stack.length) {
		const [first, last] = stack.pop()!;
		let farthest = -1,
			largest = tolerance;
		for (let index = first + 1; index < last; index++) {
			const gap = visioSegmentDistance(points[index]!, points[first]!, points[last]!);
			if (gap > largest) {
				largest = gap;
				farthest = index;
			}
		}
		if (farthest < 0) continue;
		keep[farthest] = 1;
		stack.push([first, farthest], [farthest, last]);
	}
	const result: number[] = [];
	keep.forEach((flag, index) => {
		if (flag) result.push(index);
	});
	return result;
}

export function visioSimplifyPath(
	points: readonly VisioPathPoint[],
	tolerance: number,
): VisioPathPoint[] {
	const clean = visioCleanSamples(points);
	return visioSimplifyIndices(clean, tolerance).map((index) => clean[index]!);
}

/** Smooth cubic segments through the points (uniform Catmull-Rom converted to Bezier). */
export function visioCatmullRomSegments(
	points: readonly VisioPathPoint[],
	closed: boolean,
): VisioPathSegment[] {
	const count = points.length;
	if (count < 2) return [];
	if (count === 2 && !closed) return [{ kind: 'line', x: points[1]!.x, y: points[1]!.y }];
	const at = (index: number) =>
		closed ? points[(index + count) % count]! : points[Math.max(0, Math.min(count - 1, index))]!;
	const segments: VisioPathSegment[] = [];
	for (let index = 0; index < (closed ? count : count - 1); index++) {
		const p0 = at(index - 1),
			p1 = at(index),
			p2 = at(index + 1),
			p3 = at(index + 2);
		segments.push({
			kind: 'cubic',
			x1: p1.x + (p2.x - p0.x) / 6,
			y1: p1.y + (p2.y - p0.y) / 6,
			x2: p2.x - (p3.x - p1.x) / 6,
			y2: p2.y - (p3.y - p1.y) / 6,
			x: p2.x,
			y: p2.y,
		});
	}
	return segments;
}

/** The circle through three points, or undefined when they are (nearly) collinear. */
export function visioCircleThrough(
	a: VisioPathPoint,
	b: VisioPathPoint,
	c: VisioPathPoint,
): { x: number; y: number; radius: number } | undefined {
	const d = 2 * (a.x * (b.y - c.y) + b.x * (c.y - a.y) + c.x * (a.y - b.y));
	const scale = Math.max(distance(a, b), distance(b, c), distance(a, c));
	if (!(scale > 0) || Math.abs(d) < 1e-9 * scale * scale) return undefined;
	const a2 = a.x * a.x + a.y * a.y,
		b2 = b.x * b.x + b.y * b.y,
		c2 = c.x * c.x + c.y * c.y;
	const x = (a2 * (b.y - c.y) + b2 * (c.y - a.y) + c2 * (a.y - b.y)) / d;
	const y = (a2 * (c.x - b.x) + b2 * (a.x - c.x) + c2 * (b.x - a.x)) / d;
	return { x, y, radius: distance({ x, y }, a) };
}

/** Points along an arc segment, starting after `start` and ending at the segment end. */
export function visioArcSamples(
	start: VisioPathPoint,
	segment: Extract<VisioPathSegment, { kind: 'arc' }>,
	steps = 32,
): VisioPathPoint[] {
	// Scaling y by rx / ry turns the axis-aligned ellipse into a circle.
	const k = segment.ratio;
	const s = { x: start.x, y: start.y * k },
		m = { x: segment.a, y: segment.b * k },
		e = { x: segment.x, y: segment.y * k };
	const circle = visioCircleThrough(s, m, e);
	if (!circle) return [{ x: segment.x, y: segment.y }];
	const angle = (point: VisioPathPoint) => Math.atan2(point.y - circle.y, point.x - circle.x);
	const turn = (from: number, to: number) => {
		const value = (to - from) % (2 * Math.PI);
		return value < 0 ? value + 2 * Math.PI : value;
	};
	const from = angle(s);
	// Counter-clockwise sweep when the mid point lies on the counter-clockwise path to the end.
	const ccw = turn(from, angle(m)) < turn(from, angle(e));
	const sweep = ccw ? turn(from, angle(e)) : -turn(angle(e), from);
	const result: VisioPathPoint[] = [];
	for (let step = 1; step <= steps; step++) {
		const theta = from + (sweep * step) / steps;
		result.push(
			step === steps
				? { x: segment.x, y: segment.y }
				: {
						x: circle.x + circle.radius * Math.cos(theta),
						y: (circle.y + circle.radius * Math.sin(theta)) / k,
					},
		);
	}
	return result;
}

/** Sample every segment into a polyline (start included). */
export function visioPathSamples(
	start: VisioPathPoint,
	segments: readonly VisioPathSegment[],
	steps = 16,
): VisioPathPoint[] {
	const result: VisioPathPoint[] = [{ x: start.x, y: start.y }];
	let current = start;
	for (const segment of segments) {
		if (segment.kind === 'line') result.push({ x: segment.x, y: segment.y });
		else if (segment.kind === 'arc') result.push(...visioArcSamples(current, segment, steps * 2));
		else
			for (let step = 1; step <= steps; step++) {
				const t = step / steps,
					u = 1 - t;
				const blend = (p0: number, p1: number, p2: number, p3: number) =>
					u * u * u * p0 + 3 * u * u * t * p1 + 3 * u * t * t * p2 + t * t * t * p3;
				result.push({
					x: blend(current.x, segment.x1, segment.x2, segment.x),
					y: blend(current.y, segment.y1, segment.y2, segment.y),
				});
			}
		current = { x: segment.x, y: segment.y };
	}
	return result;
}

/** Bounding box of the drawn path (sampled for curves). */
export function visioPathBounds(
	start: VisioPathPoint,
	segments: readonly VisioPathSegment[],
): { minX: number; minY: number; maxX: number; maxY: number } {
	const samples = visioPathSamples(start, segments, 32);
	return {
		minX: Math.min(...samples.map((point) => point.x)),
		minY: Math.min(...samples.map((point) => point.y)),
		maxX: Math.max(...samples.map((point) => point.x)),
		maxY: Math.max(...samples.map((point) => point.y)),
	};
}

/** Freeform: simplify the samples, then fit smooth cubic segments. */
export function visioFitFreeform(
	samples: readonly VisioPathPoint[],
	tolerance: number,
	closed: boolean,
): VisioPathSegment[] {
	const points = visioSimplifyPath(samples, tolerance);
	if (closed && points.length > 2 && distance(points[0]!, points[points.length - 1]!) < 1e-9)
		points.pop();
	return visioCatmullRomSegments(points, closed && points.length > 2);
}

function turning(a: VisioPathPoint, b: VisioPathPoint, c: VisioPathPoint): number {
	const first = Math.atan2(b.y - a.y, b.x - a.x),
		second = Math.atan2(c.y - b.y, c.x - b.x);
	const value = Math.abs(second - first) % (2 * Math.PI);
	return value > Math.PI ? 2 * Math.PI - value : value;
}

/**
 * Pencil: split the stroke at sharp corners, then emit a straight line for each nearly straight
 * piece and a circular arc for each bowed piece (S-shaped pieces split until they fit).
 */
export function visioFitPencil(
	samples: readonly VisioPathPoint[],
	tolerance: number,
): VisioPathSegment[] {
	const points = visioCleanSamples(samples);
	if (points.length < 2) return [];
	const fine = visioSimplifyIndices(points, tolerance / 2);
	const corners = [fine[0]!];
	for (let index = 1; index < fine.length - 1; index++)
		if (turning(points[fine[index - 1]!]!, points[fine[index]!]!, points[fine[index + 1]!]!) > 1)
			corners.push(fine[index]!);
	corners.push(fine[fine.length - 1]!);
	const segments: VisioPathSegment[] = [];
	const piece = (first: number, last: number, depth: number): void => {
		const a = points[first]!,
			b = points[last]!;
		let farthest = first,
			largest = 0;
		for (let index = first + 1; index < last; index++) {
			const gap = visioSegmentDistance(points[index]!, a, b);
			if (gap > largest) {
				largest = gap;
				farthest = index;
			}
		}
		if (largest <= tolerance || last - first < 2) {
			segments.push({ kind: 'line', x: b.x, y: b.y });
			return;
		}
		const middle = points[farthest]!;
		const circle = visioCircleThrough(a, middle, b);
		const fits =
			circle &&
			points
				.slice(first, last + 1)
				.every((point) => Math.abs(distance(point, circle) - circle.radius) <= tolerance * 2);
		if (fits || depth >= 6) {
			if (circle)
				segments.push({ kind: 'arc', x: b.x, y: b.y, a: middle.x, b: middle.y, ratio: 1 });
			else segments.push({ kind: 'line', x: b.x, y: b.y });
			return;
		}
		piece(first, farthest, depth + 1);
		piece(farthest, last, depth + 1);
	};
	for (let index = 1; index < corners.length; index++)
		piece(corners[index - 1]!, corners[index]!, 0);
	return segments;
}

/** A quarter of the axis-aligned ellipse whose corner opposite the centre is the drag box. */
export function visioQuarterArc(
	start: VisioPathPoint,
	end: VisioPathPoint,
): Extract<VisioPathSegment, { kind: 'arc' }> | undefined {
	const rx = Math.abs(end.x - start.x),
		ry = Math.abs(end.y - start.y);
	if (!(rx > 0) || !(ry > 0)) return undefined;
	// Centre below/above the start, level with the end: the arc leaves the start horizontally.
	const cx = start.x,
		cy = end.y,
		c = Math.SQRT1_2;
	return {
		kind: 'arc',
		x: end.x,
		y: end.y,
		a: cx + (end.x - cx) * c,
		b: cy + (start.y - cy) * c,
		ratio: rx / ry,
	};
}
