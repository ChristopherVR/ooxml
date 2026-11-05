/**
 * A shape whose flip changes between the slides TURNS over in 3D, the way
 * PowerPoint's Morph plays it.
 *
 * Measured on PowerPoint 16 (`CreateVideo`, 60 fps, 720p, Morph by object, 2 s):
 * a filled right triangle with a 6 pt outline, flipped horizontally (or
 * vertically) and moved, never squashes flat through zero. Every frame is the
 * triangle rotated about its own vertical (horizontal) centre line and seen
 * through one perspective camera:
 *
 *  - the vanishing point is the SLIDE's centre, not the shape's (fitting the
 *    shape's centre left 4.3 px of error per vertex, the slide's 0.8 px);
 *  - the eye sits 1100 px in front of a 1280 x 720 slide, which is a 60 degree
 *    horizontal field of view: `distance = (width / 2) / tan(30 deg)`
 *    (1000 px fits to 1.2 px, 1100 px to 0.8 px, 1200 px to 1.35 px, for both
 *    the horizontal and the vertical flip);
 *  - the angle runs 0 -> 180 degrees on the travel's own curve (48.0 degrees
 *    where the centre has travelled 26.6%, 122.7 where it has travelled
 *    68.3%), and the turn is always the one that brings the shape's own +x
 *    side (+y for a vertical flip, in its unflipped geometry) towards the
 *    viewer, whichever way the flip goes: flipping a right triangle swings its
 *    right side out of the slide, flipping it back swings its (mirrored) left.
 *
 * The journey's two keyframes become `translate(v) perspective(d) translate(t - v)
 * rotate(r) rotateY(a) rotateX(b) scale(sx, sy) scaleX(fx) scaleY(fy)`, with `v`
 * the slide centre seen from the element's own centre (its transform origin)
 * and `fx`, `fy` the OUTGOING flips on both frames: the turn makes the mirror,
 * so `rotateY(180deg) scaleX(fx)` paints exactly `scaleX(-fx)` at the end. Every
 * argument is linear in the eased progress, so the browser's numeric
 * per-function interpolation between two frames is the measured motion.
 *
 * @module render/morph-flip-turn
 */
import type { PptxElement } from 'ooxml-core/pptx';

/** The slide the morph plays on, in the element coordinates (px). */
export interface MorphViewSize {
	readonly width: number;
	readonly height: number;
}

/** PowerPoint's 16:9 slide in the engine's pixels, for callers that pass none. */
export const DEFAULT_MORPH_VIEW: MorphViewSize = { width: 1280, height: 720 };

/** `1 / tan(30 deg)`: the eye distance per half slide width (60 degree field of view). */
const EYE_PER_HALF_WIDTH = 1 / Math.tan(Math.PI / 6);

const round = (value: number): number => Math.round(value * 1000) / 1000;

/** Whether a pair's horizontal or vertical flip changes, so its body turns over. */
export function morphFlipTurns(from: PptxElement, to: PptxElement): boolean {
	return (
		Boolean(from.flipHorizontal) !== Boolean(to.flipHorizontal) ||
		Boolean(from.flipVertical) !== Boolean(to.flipVertical)
	);
}

/** The two journey transforms of one half of a turning pair. */
export interface MorphFlipTurnFrames {
	readonly start: string;
	readonly end: string;
}

/**
 * The start and end `transform` of the half painted as `local` (its box is the
 * frame the transform acts in, about its centre), for a pair travelling from
 * `from`'s box at `fromRot` degrees to `to`'s box at `toRot`.
 */
export function buildMorphFlipTurn(
	local: PptxElement,
	from: PptxElement,
	to: PptxElement,
	fromRot: number,
	toRot: number,
	view: MorphViewSize = DEFAULT_MORPH_VIEW,
): MorphFlipTurnFrames {
	const cx = local.x + local.width / 2;
	const cy = local.y + local.height / 2;
	const vx = view.width / 2 - cx;
	const vy = view.height / 2 - cy;
	const eye = (view.width / 2) * EYE_PER_HALF_WIDTH;
	const fx = from.flipHorizontal ? -1 : 1;
	const fy = from.flipVertical ? -1 : 1;
	const turnY = Boolean(from.flipHorizontal) !== Boolean(to.flipHorizontal) ? -fx * 180 : 0;
	const turnX = Boolean(from.flipVertical) !== Boolean(to.flipVertical) ? fy * 180 : 0;
	const frame = (box: PptxElement, rotation: number, end: boolean): string => {
		const tx = box.x + box.width / 2 - cx;
		const ty = box.y + box.height / 2 - cy;
		const sx = Math.max(box.width, 1) / Math.max(local.width, 1);
		const sy = Math.max(box.height, 1) / Math.max(local.height, 1);
		return [
			`translate(${round(vx)}px, ${round(vy)}px)`,
			`perspective(${round(eye)}px)`,
			`translate(${round(tx - vx)}px, ${round(ty - vy)}px)`,
			`rotate(${round(rotation)}deg)`,
			`rotateY(${end ? turnY : 0}deg)`,
			`rotateX(${end ? turnX : 0}deg)`,
			`scale(${round(sx)}, ${round(sy)})`,
			`scaleX(${fx}) scaleY(${fy})`,
		].join(' ');
	};
	return { start: frame(from, fromRot, false), end: frame(to, toRot, true) };
}
