/**
 * How far ahead of its eased progress PowerPoint's Morph grows a line's box.
 *
 * Measured on PowerPoint 16 (`CreateVideo`, 60 fps, 720p, Morph by object): a
 * straight or elbow connector's centre follows the eased progress `p` exactly,
 * but each extent (width, height) overshoots the linear blend of its two ends.
 * Per axis, with `lo`, `hi` the smaller and larger extent of the pair, `r = hi / lo`:
 *
 *   extent(p) = lerp(a0, a1, p) + |a1 - a0| * hi / (hi + lo) * G * p * (1 - p)
 *   G = (1 - exp(-3.2 (r - 1))) * coupling(x)
 *
 * so a grow AND a shrink are both larger than the straight blend mid-flight
 * (3x: +0.75 * |delta| * p(1-p), 4x: +0.80, 10x: +0.91). The ratio term is the
 * small-ratio fade: a 1.1x resize has no overshoot, 1.25x half of it, 1.4x 70%,
 * 2x 95% (measured at 1.1x to 2x, grow and shrink: the old law without it was 4 pt
 * off at 1.4x). `coupling` ties the resize to the centre's travel `d` along the same
 * axis, with `x = d / (|a1 - a0| / 2)` (`x = 1` is an end that stays put): `x^1.14`
 * while `x <= 1` (a box that grows in place about a fixed centre, `x = 0`, is exactly
 * linear), and `min(1, 1.1 / x^0.94)` beyond. Fitted to 43 synthetic cases: mean abs
 * error 1.5 pt where the straight blend was 40 pt.
 *
 * The law runs on the SIGNED extent along each slide axis, which makes it cover:
 *
 *  - rotated lines: PowerPoint blends the two ends' END POINTS (a pure rotation about
 *    a fixed centre shortens the chord mid-turn, as the blend of the end points does;
 *    a resize about a fixed centre has no overshoot), then the overshoot above applies
 *    to each slide axis' span with the centre travel taken per axis. That makes a
 *    line that turns while its centre moves 112 pt measure 306 long mid-turn (the bare
 *    chord is 272), as PowerPoint does. Two ends that share one rotation (a line that
 *    keeps its angle and resizes) overshoot along their own length instead, measured
 *    at 2x to within 1.4 pt;
 *  - an axis the two ends flip differently on: the extent passes through zero, and
 *    the overshoot is `hi * (r - 1) / (r - 1 + 0.56) * coupling(x)` (x from the
 *    magnitudes) towards the larger end, where `hi` is the larger magnitude
 *    (1.5x: 0.47, 2x: 0.65, 4x: 0.81 of `hi`; equal magnitudes: none).
 *
 * Freeform open paths and arcs were measured too and do NOT overshoot: their box
 * follows the plain linear blend (fraction of the resize done equals `p` to 0.01).
 *
 * Plain rectangles do not show it, so only lines and connectors use this.
 *
 * @module render/morph-line-reach
 */
import type { PptxElement } from 'ooxml-core/pptx';

const RATE = 3.2;
const FLIP_HALF_SATURATION = 0.56;

/** The coupling of a resize to the centre's travel (`x` in half-deltas). */
export function reachCoupling(x: number): number {
	return x <= 1 ? x ** 1.14 : Math.min(1, 1.1 / x ** 0.94);
}

/**
 * One axis' SIGNED extent overshoot at eased progress `p`: what PowerPoint adds to
 * the linear blend of the signed extents `s0` and `s1` (positive towards the larger
 * end's sign). `centreTravel` is the centre's absolute move along the axis.
 */
export function reachOffset(s0: number, s1: number, centreTravel: number, p: number): number {
	const a0 = Math.abs(s0);
	const a1 = Math.abs(s1);
	const hi = Math.max(a0, a1);
	const lo = Math.min(a0, a1);
	const delta = hi - lo;
	if (hi < 1e-6 || delta < 1e-6) return 0;
	const direction = Math.sign(a0 >= a1 ? s0 : s1);
	const coupling = reachCoupling(centreTravel / (delta / 2));
	const weight = p * (1 - p);
	if (s0 * s1 < 0) {
		const r = hi / lo;
		return direction * hi * ((r - 1) / (r - 1 + FLIP_HALF_SATURATION)) * coupling * weight;
	}
	const fade = lo < 1e-6 ? 1 : 1 - Math.exp(-RATE * (hi / lo - 1));
	return direction * delta * (hi / (hi + lo)) * fade * coupling * weight;
}

/**
 * One axis' magnitude at eased progress `p`: the linear blend plus PowerPoint's
 * overshoot. `centreTravel` is the centre's absolute move along the axis.
 */
export function reachedExtent(a0: number, a1: number, centreTravel: number, p: number): number {
	const linear = a0 + (a1 - a0) * p;
	if (a0 < 0 || a1 < 0) return linear;
	return linear + reachOffset(a0, a1, centreTravel, p);
}

/** A box's end-to-end vector on the slide (flips and rotation applied). */
function boxVector(el: PptxElement, rotated: boolean): readonly [number, number] {
	const vx = (el.flipHorizontal ? -1 : 1) * el.width;
	const vy = (el.flipVertical ? -1 : 1) * el.height;
	if (!rotated) return [vx, vy];
	const turn = ((el.rotation ?? 0) * Math.PI) / 180;
	const [cos, sin] = [Math.cos(turn), Math.sin(turn)];
	return [cos * vx - sin * vy, sin * vx + cos * vy];
}

/**
 * What PowerPoint adds to the linearly blended end-to-end vector of a pair at `p`,
 * `[dx, dy]` on the slide axes. `allowRotation` is true for a straight line (its end
 * points are blended whatever the rotation and flips); an elbow or curved connector
 * whose box turns is blended on the box instead, and one that flips on an axis was
 * not measured either, so those get none.
 */
export function reachOffsets(
	from: PptxElement,
	to: PptxElement,
	p: number,
	allowRotation = true,
): readonly [number, number] {
	const rotFrom = from.rotation ?? 0;
	const rotTo = to.rotation ?? 0;
	if (!allowRotation && (rotFrom !== 0 || rotTo !== 0)) return [0, 0];
	const [fx, fy] = boxVector(from, true);
	const [tx, ty] = boxVector(to, true);
	const dx = Math.abs(to.x + to.width / 2 - (from.x + from.width / 2));
	const dy = Math.abs(to.y + to.height / 2 - (from.y + from.height / 2));
	if (rotFrom === rotTo && rotFrom !== 0) {
		// Two ends turned alike resize along their own length (a line that keeps its
		// angle and doubles overshoots by the unrotated law on its length).
		const length = reachOffset(Math.hypot(fx, fy), Math.hypot(tx, ty), Math.hypot(dx, dy), p);
		const [bx, by] = [fx + (tx - fx) * p, fy + (ty - fy) * p];
		const chord = Math.hypot(bx, by);
		return chord < 1e-6 ? [0, 0] : [(bx / chord) * length, (by / chord) * length];
	}
	// Otherwise the end points are blended and each slide axis overshoots on its own. Only
	// straight lines were measured through a flip; an elbow's axis stays linear there.
	const axis = (s0: number, s1: number, travel: number): number =>
		!allowRotation && s0 * s1 < 0 ? 0 : reachOffset(s0, s1, travel, p);
	return [axis(fx, tx, dx), axis(fy, ty, dy)];
}

/**
 * The factors by which a pair's linearly blended width and height are scaled at
 * `p` (1 when an axis has no overshoot), for the pair's unrotated, unflipped-alike
 * axes. Kept for callers that scale rather than offset.
 */
export function reachFactors(
	from: PptxElement,
	to: PptxElement,
	p: number,
): readonly [number, number] {
	const [dx, dy] = reachOffsets(from, to, p, false);
	const [fx, fy] = boxVector(from, false);
	const [tx, ty] = boxVector(to, false);
	const scale = (offset: number, s0: number, s1: number): number => {
		const linear = s0 + (s1 - s0) * p;
		return Math.abs(linear) < 1e-6 ? 1 : (linear + offset) / linear;
	};
	return [scale(dx, fx, tx), scale(dy, fy, ty)];
}
