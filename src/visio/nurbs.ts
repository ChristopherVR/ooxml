import { flattenNurbs, CURVE_TOLERANCE } from './nurbs-flatten.js';
import type { Report } from './sheet.js';

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
	if (initial < degree + 1) {
		report(
			'unsupported-geometry',
			'Periodic NURBS is unsupported; its geometry section was omitted.',
		);
		return;
	}
	if (initial > degree + 1) return invalid('NURBS first-knot multiplicity exceeds degree + 1');
	let repeats = 1;
	for (let i = initial; i < knots.length; i++) {
		repeats = knots[i] === knots[i - 1] ? repeats + 1 : 1;
		if (repeats > degree || (i === knots.length - 1 && repeats > 1))
			return invalid('Invalid NURBS internal or last-knot multiplicity');
	}
	// Visio stores n + 1 knots for n controls; the repeated end knots are implicit.
	for (let i = 0; i < degree; i++) knots.push(lastKnot);
	const points = flattenNurbs(controls, knots, degree, consume, consumeWork);
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
	return points.map((p) => `L ${p.x === 0 ? 0 : p.x} ${p.y === 0 ? 0 : p.y}`).join(' ');
}
