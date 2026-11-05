/**
 * How far ahead of its eased progress PowerPoint's Morph grows a line's box.
 *
 * Measured on PowerPoint 16 (`CreateVideo`, 60 fps, 720p, Morph by object): a
 * straight or elbow connector's centre follows the eased progress `p` exactly,
 * but each extent (width, height) overshoots the linear blend of its two ends.
 * Per axis, with `lo`, `hi` the smaller and larger extent of the pair:
 *
 *   extent(p) = lerp(a0, a1, p) + |a1 - a0| * hi / (hi + lo) * g * p * (1 - p)
 *
 * so a grow AND a shrink are both larger than the straight blend mid-flight
 * (3x: +0.75 * |delta| * p(1-p), 4x: +0.80, 10x: +0.91). `g` couples the
 * resize to the centre's travel `d` along the same axis, with `x = d / (|a1 - a0| / 2)`
 * (`x = 1` is an end that stays put): `g = x^1.1` while `x <= 1` (a box that
 * grows in place about a fixed centre, `x = 0`, is exactly linear), and
 * `g = min(1, 1.14 / x)` beyond. Fitted to 28 synthetic cases (several start
 * sizes and aspects, 1.4x to 10x, shrinks, independent axes, centre moves of
 * 0 to 3.5 half-deltas): mean abs error 2 pt where the straight blend was 40 pt.
 * Plain rectangles do not show it, so only lines and connectors use this.
 *
 * @module render/morph-line-reach
 */
import type { PptxElement } from 'ooxml-core/pptx';

/** The coupling of a resize to the centre's travel (`x` in half-deltas). */
export function reachCoupling(x: number): number {
	return x <= 1 ? x ** 1.1 : Math.min(1, 1.14 / x);
}

/**
 * One axis' magnitude at eased progress `p`: the linear blend plus PowerPoint's
 * overshoot. `centreTravel` is the centre's absolute move along the axis.
 */
export function reachedExtent(a0: number, a1: number, centreTravel: number, p: number): number {
	const delta = a1 - a0;
	const linear = a0 + delta * p;
	const size = Math.abs(delta);
	if (size < 1e-6) return linear;
	const lo = Math.min(a0, a1);
	const hi = Math.max(a0, a1);
	if (lo < 0 || hi <= 0) return linear;
	const coupling = reachCoupling(centreTravel / (size / 2));
	return linear + ((size * hi) / (hi + lo)) * coupling * p * (1 - p);
}

/**
 * The factors by which a pair's linearly blended width and height are scaled at
 * `p` (1 when an axis has no overshoot). Rotated pairs are left linear (the
 * measurements are unrotated), and so is an axis the two ends flip differently on (its extent then
 * passes through zero or turns, which the measurements do not cover).
 */
export function reachFactors(
	from: PptxElement,
	to: PptxElement,
	p: number,
): readonly [number, number] {
	if ((from.rotation ?? 0) !== 0 || (to.rotation ?? 0) !== 0) return [1, 1];
	const axis = (
		a0: number,
		a1: number,
		c0: number,
		c1: number,
		flip0: boolean | undefined,
		flip1: boolean | undefined,
	): number => {
		if (Boolean(flip0) !== Boolean(flip1)) return 1;
		const linear = a0 + (a1 - a0) * p;
		if (linear < 1e-6) return 1;
		return reachedExtent(a0, a1, Math.abs(c1 - c0), p) / linear;
	};
	return [
		axis(
			from.width,
			to.width,
			from.x + from.width / 2,
			to.x + to.width / 2,
			from.flipHorizontal,
			to.flipHorizontal,
		),
		axis(
			from.height,
			to.height,
			from.y + from.height / 2,
			to.y + to.height / 2,
			from.flipVertical,
			to.flipVertical,
		),
	];
}
