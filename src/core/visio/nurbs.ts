import { flattenNurbs, CURVE_TOLERANCE } from './nurbs-flatten';
import type { Report } from './sheet';

export type NurbsPoint = { x: number; y: number };
export type NurbsControl = NurbsPoint & { weight: number };
type Point = NurbsPoint;
const bounded = (n: number) => Number.isFinite(n) && Math.abs(n) <= 1e9;
const validPoint = (p: Point) => bounded(p.x) && bounded(p.y);
/** Validate and expand Visio's compact n+1 knot sequence before shared bounded sampling. */
export function normalizedNurbs(
	controls: NurbsControl[],
	compactKnots: number[],
	degree: number,
	report: Report,
	consume: () => void,
	kind = 'NURBSTo',
	consumeWork: (units: number) => void = () => {},
): string | undefined {
	const invalid = (message: string) => {
		report('invalid-geometry', `${message}; its geometry section was omitted.`);
		return undefined;
	};
	if (
		!Number.isInteger(degree) ||
		degree < 1 ||
		degree > 25 ||
		controls.length <= degree ||
		controls.length > 256
	)
		return invalid(
			'NURBS requires degree 1 to 25 and 2 to 256 control points, more than its degree',
		);
	if (controls.some((p) => !validPoint(p) || !bounded(p.weight) || p.weight < 1e-9))
		return invalid('NURBS supports bounded coordinates and positive weights from 1e-9 to 1e9');
	if (
		compactKnots.length !== controls.length + 1 ||
		compactKnots.some((k, i) => !bounded(k) || (i > 0 && k < compactKnots[i - 1]!)) ||
		compactKnots[0] === compactKnots.at(-1)
	)
		return invalid('NURBS knots must be nondecreasing with a nonempty parameter range');
	const knots = [...compactKnots],
		lastKnot = knots.at(-1)!;
	let initial = 1;
	while (initial < knots.length && knots[initial] === knots[0]) initial++;
	if (initial > degree + 1) return invalid('NURBS first-knot multiplicity exceeds degree + 1');
	const periodic = initial < degree + 1;
	let repeats = 1;
	for (let i = initial; i < knots.length; i++) {
		repeats = knots[i] === knots[i - 1] ? repeats + 1 : 1;
		if (repeats > degree || (i === knots.length - 1 && repeats > 1))
			return invalid('Invalid NURBS internal or last-knot multiplicity');
	}
	let evalControls: readonly NurbsControl[] = controls;
	let evalKnots: number[];
	if (periodic) {
		// Periodic Visio NURBS wraps the first `degree` controls and extends the n + 1
		// knot sequence periodically across period = lastKnot - firstKnot (shifted to 0).
		const count = controls.length;
		const base = knots[0]!;
		const period = lastKnot - base;
		evalControls = [...controls, ...Array.from({ length: degree }, (_, i) => controls[i % count]!)];
		evalKnots = Array.from({ length: count + 2 * degree + 1 }, (_, i) => {
			const cycles = Math.floor(i / count);
			return cycles * period + (knots[i % count]! - base);
		});
	} else {
		// Visio stores n + 1 knots for n controls; the repeated end knots are implicit.
		for (let i = 0; i < degree; i++) knots.push(lastKnot);
		evalKnots = knots;
	}
	const points = flattenNurbs(evalControls, evalKnots, degree, consume, consumeWork, periodic);
	if (!points) {
		report(
			'geometry-limit',
			'NURBS subdivision exceeded its numeric, depth, segment or work bound; its geometry section was omitted.',
		);
		return;
	}
	report(
		'geometry-approximation',
		`${kind} was approximated by line segments with ${CURVE_TOLERANCE}-inch control-hull flatness; floating-point curve parity is not certified.`,
	);
	const fmt = (p: Point) => `${p.x === 0 ? 0 : p.x} ${p.y === 0 ? 0 : p.y}`;
	if (periodic) {
		const [first, ...rest] = points;
		return `M ${fmt(first!)} ${rest.map((p) => `L ${fmt(p)}`).join(' ')} Z`;
	}
	return points.map((p) => `L ${fmt(p)}`).join(' ');
}
