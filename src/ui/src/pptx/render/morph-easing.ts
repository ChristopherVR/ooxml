/**
 * Cubic-bezier arithmetic for the morph timing curves.
 *
 * CSS can only state a timing function per keyframe interval, while PowerPoint
 * drives every part of a morph (travel, resize, rotation, fill colour, and the
 * dissolve of unmatched shapes) from one eased progress value. These helpers
 * let a keyframe interval reproduce an exact slice of that one curve, so the
 * dissolve windows are stated in eased progress and still render with plain
 * CSS keyframes.
 */

/** Control points of a CSS `cubic-bezier(x1, y1, x2, y2)`. */
export type BezierPoints = readonly [number, number, number, number];

/** One axis of a cubic Bezier whose control values are 0, a, b, 1. */
function axisAt(a: number, b: number, u: number): number {
	return 3 * a * u * (1 - u) * (1 - u) + 3 * b * u * u * (1 - u) + u * u * u;
}

function axisDerivative(a: number, b: number, u: number): number {
	return 3 * (1 - u) * (1 - u) * a + 6 * (1 - u) * u * (b - a) + 3 * u * u * (1 - b);
}

/**
 * The curve parameter `u` at which one axis reaches `value` (0..1). Newton
 * where the derivative is usable, bisection where it is not; the axis must be
 * monotone, which holds for CSS time axes and for the progress axis of every
 * morph curve here.
 */
function parameterAt(a: number, b: number, value: number): number {
	if (value <= 0) {
		return 0;
	}
	if (value >= 1) {
		return 1;
	}
	let u = value;
	for (let i = 0; i < 8; i++) {
		const err = axisAt(a, b, u) - value;
		if (Math.abs(err) < 1e-7) {
			return u;
		}
		const slope = axisDerivative(a, b, u);
		if (Math.abs(slope) < 1e-6) {
			break;
		}
		u -= err / slope;
		if (u <= 0 || u >= 1) {
			break;
		}
	}
	let lo = 0;
	let hi = 1;
	u = value;
	for (let i = 0; i < 40; i++) {
		if (axisAt(a, b, u) < value) {
			lo = u;
		} else {
			hi = u;
		}
		u = (lo + hi) / 2;
	}
	return u;
}

/** Progress along a CSS cubic-bezier easing at normalised time `t` (0..1). */
export function easedProgress(t: number, x1: number, y1: number, x2: number, y2: number): number {
	return axisAt(y1, y2, parameterAt(x1, x2, t));
}

/** The normalised time (0..1) at which the easing reaches `progress` (0..1). */
export function timeAtProgress(progress: number, points: BezierPoints): number {
	const [x1, y1, x2, y2] = points;
	return axisAt(x1, x2, parameterAt(y1, y2, progress));
}

/** A slice of an easing curve, ready to be one keyframe interval. */
export interface EasingSegment {
	/** Where the slice starts, in normalised time (0..1). */
	readonly startTime: number;
	/** Where the slice ends, in normalised time (0..1). */
	readonly endTime: number;
	/** The slice renormalised to a unit square, as a CSS timing function. */
	readonly easing: string;
}

function round(value: number): number {
	return Math.round(value * 10000) / 10000;
}

/**
 * The part of `points` between eased progress `fromProgress` and `toProgress`,
 * as a keyframe interval and the CSS timing function that reproduces it.
 *
 * A property interpolated linearly from one keyframe to the next on this
 * timing function moves in exact proportion to the whole curve's progress, so
 * "linear in eased progress between a and b" becomes two keyframes and one
 * `animation-timing-function`. The slice of a cubic Bezier is itself a cubic
 * Bezier (de Casteljau), so nothing is approximated.
 */
export function easingSegment(
	points: BezierPoints,
	fromProgress: number,
	toProgress: number,
): EasingSegment {
	const [x1, y1, x2, y2] = points;
	const u0 = parameterAt(y1, y2, fromProgress);
	const u1 = parameterAt(y1, y2, toProgress);
	// Control polygon of the slice [u0, u1] of one axis with values 0, a, b, 1.
	const slice = (a: number, b: number): [number, number, number, number] => {
		const at = (u: number) => axisAt(a, b, u);
		const d = (u: number) => axisDerivative(a, b, u);
		const span = u1 - u0;
		const p0 = at(u0);
		const p3 = at(u1);
		return [p0, p0 + (d(u0) * span) / 3, p3 - (d(u1) * span) / 3, p3];
	};
	const [tx0, tx1, tx2, tx3] = slice(x1, x2);
	const [py0, py1, py2, py3] = slice(y1, y2);
	const tw = tx3 - tx0 || 1;
	const ph = py3 - py0 || 1;
	// CSS rejects a time control point outside [0, 1]; a monotone time axis
	// keeps them inside, the clamp only absorbs floating-point noise.
	const time = (v: number) => round(Math.min(1, Math.max(0, (v - tx0) / tw)));
	const easing = `cubic-bezier(${time(tx1)}, ${round((py1 - py0) / ph)}, ${time(tx2)}, ${round(
		(py2 - py0) / ph,
	)})`;
	return { startTime: tx0, endTime: tx3, easing };
}
