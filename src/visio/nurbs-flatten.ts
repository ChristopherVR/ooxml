import type { NurbsControl, NurbsPoint } from './nurbs.js';

// Knot insertion preserves the curve. Positive-weight rational Bezier curves stay
// inside the Euclidean control hull, so hull-to-segment distance bounds deviation.
// https://pages.mtu.edu/~shene/COURSES/cs3621/NOTES/spline/NURBS-knot-insert.html
// https://pages.mtu.edu/~shene/COURSES/cs3621/NOTES/spline/B-spline/bspline-curve-prop.html
// This remains floating-point approximation, not a numerical certification.
type Homogeneous = readonly [number, number, number];
export const CURVE_TOLERANCE = 0.0001;
const MAX_WORK = 2_000_000;
const MAX_SEGMENTS = 2048;
const MAX_REFINED_CONTROLS = 6400;
const MAX_DEPTH = 20;
const failure = Symbol('curve-limit');
const project = (p: Homogeneous): NurbsPoint => ({ x: p[0] / p[2], y: p[1] / p[2] });
function blend(a: Homogeneous, b: Homogeneous, t: number): Homogeneous {
	return [(1 - t) * a[0] + t * b[0], (1 - t) * a[1] + t * b[1], (1 - t) * a[2] + t * b[2]];
}
function distance(p: NurbsPoint, a: NurbsPoint, b: NurbsPoint): number {
	const dx = b.x - a.x,
		dy = b.y - a.y,
		squared = dx * dx + dy * dy;
	const t = squared ? Math.max(0, Math.min(1, ((p.x - a.x) * dx + (p.y - a.y) * dy) / squared)) : 0;
	return Math.hypot(p.x - a.x - t * dx, p.y - a.y - t * dy);
}
/** Private numeric kernel. Inputs must already satisfy normalizedNurbs validation. */
export function flattenNurbs(
	controls: readonly NurbsControl[],
	sourceKnots: readonly number[],
	degree: number,
	consume: () => void,
	consumeWork: (units: number) => void,
): NurbsPoint[] | undefined {
	let work = 0;
	const charge = (units: number) => {
		consumeWork(units);
		if ((work += units) > MAX_WORK) throw failure;
	};
	try {
		charge(controls.length + sourceKnots.length);
		let points: Homogeneous[] = controls.map((p) => [p.x * p.weight, p.y * p.weight, p.weight]);
		let knots = [...sourceKnots];
		// Expand each internal knot to degree multiplicity. Refined spans are Beziers.
		const internal = [...new Set(knots.slice(degree + 1, -degree - 1))];
		for (const u of internal) {
			charge(knots.length);
			let multiplicity = knots.filter((k) => k === u).length;
			while (multiplicity < degree) {
				charge(points.length + knots.length + degree * 4);
				if (points.length >= MAX_REFINED_CONTROLS) throw failure;
				let span = degree;
				while (span < points.length - 1 && knots[span + 1]! <= u) span++;
				const next: Homogeneous[] = new Array(points.length + 1);
				for (let i = 0; i <= span - degree; i++) next[i] = points[i]!;
				for (let i = span - multiplicity; i < points.length; i++) next[i + 1] = points[i]!;
				for (let i = span - degree + 1; i <= span - multiplicity; i++) {
					const alpha = (u - knots[i]!) / (knots[i + degree]! - knots[i]!);
					if (!Number.isFinite(alpha) || alpha < 0 || alpha > 1) throw failure;
					next[i] = blend(points[i - 1]!, points[i]!, alpha);
				}
				knots = [...knots.slice(0, span + 1), u, ...knots.slice(span + 1)];
				points = next;
				multiplicity++;
			}
		}
		const output: NurbsPoint[] = [];
		const subdivide = (h: readonly Homogeneous[], depth: number): void => {
			charge(h.length);
			const p = h.map(project),
				a = p[0]!,
				b = p.at(-1)!;
			if (
				p.some(
					(q) =>
						!Number.isFinite(q.x) ||
						!Number.isFinite(q.y) ||
						Math.abs(q.x) > 1e9 ||
						Math.abs(q.y) > 1e9,
				)
			)
				throw failure;
			const flat = p.every((q) => distance(q, a, b) <= CURVE_TOLERANCE);
			// Retain at least four segments per span for consistent existing rendering.
			if (depth >= 2 && flat) {
				if (output.length >= MAX_SEGMENTS) throw failure;
				consume();
				output.push(b);
				return;
			}
			if (depth >= MAX_DEPTH) throw failure;
			charge(h.length * h.length);
			const row = [...h],
				left: Homogeneous[] = [row[0]!],
				right: Homogeneous[] = [row.at(-1)!];
			for (let count = row.length - 1; count > 0; count--) {
				for (let i = 0; i < count; i++) row[i] = blend(row[i]!, row[i + 1]!, 0.5);
				left.push(row[0]!);
				right.push(row[count - 1]!);
			}
			subdivide(left, depth + 1);
			subdivide(right.reverse(), depth + 1);
		};
		for (let span = degree; span < points.length; span++) {
			charge(1);
			if (knots[span + 1]! > knots[span]!) subdivide(points.slice(span - degree, span + 1), 0);
		}
		if (output.length) {
			const last = controls.at(-1)!;
			output[output.length - 1] = { x: last.x, y: last.y };
		}
		return output.length ? output : undefined;
	} catch (error) {
		if (error === failure) return undefined;
		throw error;
	}
}
