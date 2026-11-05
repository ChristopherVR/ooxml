/**
 * An SVG path `d` string as DrawingML path segments: move, line, cubic and
 * quadratic Bezier, close. Every SVG command is accepted, absolute and
 * relative (`M L H V C S Q T A Z`), with implicit repeats; `H` / `V` become
 * lines, `S` / `T` their explicit Beziers, and an elliptical arc `A` cubic
 * Beziers of at most a quarter turn each, never a polyline.
 *
 * Built on the shared path tokenizer (`svg-path-tokens.ts`) and arc centre
 * conversion (`svgArcCenter` in `svg-path-curves.ts`) that the path
 * flattener uses.
 *
 * @module geometry/svg-path-segments
 */

import { svgArcCenter, svgArcPoint } from './svg-path-curves';
import { tokenizeSvgPath } from './svg-path-tokens';
import type { Point2 } from './svg-path-types';

/** One DrawingML path command (`a:moveTo`, `a:lnTo`, `a:cubicBezTo`, `a:quadBezTo`, `a:close`). */
export type SvgPathSegment =
	| { type: 'moveTo'; pt: Point2 }
	| { type: 'lineTo'; pt: Point2 }
	| { type: 'cubicBezTo'; pts: [Point2, Point2, Point2] }
	| { type: 'quadBezTo'; pts: [Point2, Point2] }
	| { type: 'close' };

const ARG_COUNT: Readonly<Record<string, number>> = {
	M: 2,
	L: 2,
	H: 1,
	V: 1,
	C: 6,
	S: 4,
	Q: 4,
	T: 2,
	A: 7,
	Z: 0,
};

/** The number at `index`; the caller has checked there are enough. */
function num(values: readonly number[], index: number): number {
	return values[index] ?? 0;
}

function reflect(control: Point2, about: Point2): Point2 {
	return { x: 2 * about.x - control.x, y: 2 * about.y - control.y };
}

/**
 * Parse an SVG path into DrawingML path segments in the path's own
 * coordinates.
 *
 * @throws SyntaxError for path data that does not start with a command,
 *   holds a character that is neither a command nor a number, gives a
 *   command too few numbers, or holds a number that is not finite.
 */
export function svgPathToSegments(d: string): SvgPathSegment[] {
	const out: SvgPathSegment[] = [];
	let cur: Point2 = { x: 0, y: 0 };
	let start: Point2 = { x: 0, y: 0 };
	// The control point S / T reflect, and which kind of curve left it.
	let lastCubic: Point2 | undefined;
	let lastQuad: Point2 | undefined;

	for (const { type, values } of tokenizeSvgPath(d, { strict: true })) {
		const upper = type.toUpperCase();
		const relative = type !== upper;
		const count = ARG_COUNT[upper] ?? 0;
		if (count === 0) {
			out.push({ type: 'close' });
			cur = start;
			lastCubic = lastQuad = undefined;
			continue;
		}
		if (values.length === 0 || values.length % count !== 0) {
			throw new SyntaxError(`SVG path command ${type} takes ${count} numbers per segment`);
		}
		if (values.some((value) => !Number.isFinite(value))) {
			throw new SyntaxError(`SVG path command ${type} has a number that is not finite`);
		}
		for (let at = 0; at < values.length; at += count) {
			const a = values.slice(at, at + count);
			const pt = (x: number, y: number): Point2 =>
				relative ? { x: cur.x + x, y: cur.y + y } : { x, y };
			// A moveto's extra pairs are linetos.
			const command = upper === 'M' && at > 0 ? 'L' : upper;
			let nextCubic: Point2 | undefined;
			let nextQuad: Point2 | undefined;
			switch (command) {
				case 'M':
					cur = start = pt(num(a, 0), num(a, 1));
					out.push({ type: 'moveTo', pt: cur });
					break;
				case 'L':
					cur = pt(num(a, 0), num(a, 1));
					out.push({ type: 'lineTo', pt: cur });
					break;
				case 'H':
					cur = { x: relative ? cur.x + num(a, 0) : num(a, 0), y: cur.y };
					out.push({ type: 'lineTo', pt: cur });
					break;
				case 'V':
					cur = { x: cur.x, y: relative ? cur.y + num(a, 0) : num(a, 0) };
					out.push({ type: 'lineTo', pt: cur });
					break;
				case 'C':
				case 'S': {
					const rest = command === 'C' ? a.slice(2) : a;
					const c1 =
						command === 'C' ? pt(num(a, 0), num(a, 1)) : lastCubic ? reflect(lastCubic, cur) : cur;
					const c2 = pt(num(rest, 0), num(rest, 1));
					const end = pt(num(rest, 2), num(rest, 3));
					out.push({ type: 'cubicBezTo', pts: [c1, c2, end] });
					nextCubic = c2;
					cur = end;
					break;
				}
				case 'Q':
				case 'T': {
					const c =
						command === 'Q' ? pt(num(a, 0), num(a, 1)) : lastQuad ? reflect(lastQuad, cur) : cur;
					const end = command === 'Q' ? pt(num(a, 2), num(a, 3)) : pt(num(a, 0), num(a, 1));
					out.push({ type: 'quadBezTo', pts: [c, end] });
					nextQuad = c;
					cur = end;
					break;
				}
				case 'A': {
					const end = pt(num(a, 5), num(a, 6));
					const flags = [num(a, 3) !== 0, num(a, 4) !== 0] as const;
					for (const pts of arcToCubics(cur, num(a, 0), num(a, 1), num(a, 2), ...flags, end)) {
						out.push({ type: 'cubicBezTo', pts });
					}
					cur = end;
					break;
				}
			}
			lastCubic = nextCubic;
			lastQuad = nextQuad;
		}
	}
	return out;
}

/**
 * An SVG elliptical arc from `from` to `to` as cubic Beziers, each spanning
 * at most 90 degrees. A zero radius is a straight line (as SVG draws it),
 * written as a degenerate cubic; coincident end points draw nothing.
 */
export function arcToCubics(
	from: Point2,
	rx: number,
	ry: number,
	xAxisRotationDeg: number,
	largeArc: boolean,
	sweep: boolean,
	to: Point2,
): Array<[Point2, Point2, Point2]> {
	if (from.x === to.x && from.y === to.y) {
		return [];
	}
	const arc = svgArcCenter(from, rx, ry, xAxisRotationDeg, largeArc, sweep, to);
	if (!arc) {
		return [[from, to, to]];
	}
	const pieces = Math.max(1, Math.ceil(Math.abs(arc.deltaTheta) / (Math.PI / 2) - 1e-9));
	const step = arc.deltaTheta / pieces;
	const k = (4 / 3) * Math.tan(step / 4);
	// The tangent of the ellipse at `theta`, scaled by `k`, in slide space.
	const handle = (theta: number, sign: number): Point2 => {
		const x = -arc.rx * Math.sin(theta) * k * sign;
		const y = arc.ry * Math.cos(theta) * k * sign;
		return { x: x * arc.cosPhi - y * arc.sinPhi, y: x * arc.sinPhi + y * arc.cosPhi };
	};
	const out: Array<[Point2, Point2, Point2]> = [];
	let a = arc.theta1;
	for (let i = 0; i < pieces; i++) {
		const b = a + step;
		const p0 = svgArcPoint(arc, a);
		const p3 = svgArcPoint(arc, b);
		const h0 = handle(a, 1);
		const h3 = handle(b, -1);
		out.push([
			{ x: p0.x + h0.x, y: p0.y + h0.y },
			{ x: p3.x + h3.x, y: p3.y + h3.y },
			// The last end is the exact target, so rounding opens no gap.
			i === pieces - 1 ? { x: to.x, y: to.y } : p3,
		]);
		a = b;
	}
	return out;
}
