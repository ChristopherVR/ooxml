/**
 * The element transforms that carry a text mapping (`morph-text-frame`): the
 * frames a matched pair's two halves start and end on so their glyphs stay
 * laid over each other.
 *
 * Every transform is `translate(dx, dy) rotate(r) scale(sx, sy)`, followed by
 * the element's own flips and, for a TURNING pair, `translate(-tail)`, about
 * the element's own centre: the keyframe layout every morph journey shares.
 *
 * A pair standing at one rotation keeps the translation-only frames it always
 * had. A pair whose rotation CHANGES turns about a pivot instead. Measured on
 * PowerPoint 16 (`CreateVideo`, 60 fps, 720p), a 32 pt line in a box turning
 * 0 -> 40 degrees while its box widens 262 -> 520 pt and moves: the text is
 * never stretched with its box (its ink stays 324 px long on every frame),
 * and the point that travels in a STRAIGHT line is the middle of the text
 * (53% along its ink; within 0.9 px of the line at every frame, at the
 * rotation's own progress), not the box's centre, which would have swung the
 * line's start 25 px off that path. The two faces of a restyle turn the same
 * way, one glyph-aligned copy whose faces dissolve.
 *
 * The CSS function list interpolates each function on its own, so a pivot
 * that travels straight while the box turns about it is
 * `translate(A(t) - centre) rotate(r(t)) scale(s(t)) translate(-pivot)`:
 * the trailing translation lays the pivot on the element's centre, which the
 * leading one then carries along the straight path. Both halves share the
 * path, the angle and (up to the mapping's own ratio) the scale, so they stay
 * glyph over glyph on every frame.
 *
 * @module render/morph-text-transform
 */
import type { PptxElement } from 'ooxml-core/pptx';

/** A slide point, in px. */
export interface MorphPoint {
	x: number;
	y: number;
}

/**
 * How the outgoing half's text maps onto the incoming half's: a point at
 * `from + v` (slide px) corresponds to `to + scale * v`.
 */
export interface MorphTextAlignment {
	from: MorphPoint;
	to: MorphPoint;
	scaleX: number;
	scaleY: number;
	/**
	 * The rotation (degrees) BOTH halves stand at, when it is not zero. The
	 * anchors and scales are then in the boxes' own, unrotated coordinates, and
	 * `centres` holds the two boxes' centres the rotation turns them about.
	 */
	rotation?: number;
	centres?: { from: MorphPoint; to: MorphPoint };
	/**
	 * Set when the halves stand at DIFFERENT rotations: each half's rotation
	 * and the pivot (in its own unrotated slide frame) the text turns about.
	 * `centres` is then always set. `pivots.to` is the image of `pivots.from`
	 * under the mapping.
	 */
	turn?: { from: number; to: number; pivots: { from: MorphPoint; to: MorphPoint } };
}

/**
 * A `translate(dx, dy) scale(sx, sy)` about the element's own centre, then
 * (for a turning pair) `translate(-tail)`: `tail` is the pivot's offset from
 * the element's centre in its unrotated, unflipped-as-painted frame.
 */
export interface MorphTextTransform {
	dx: number;
	dy: number;
	sx: number;
	sy: number;
	tail?: MorphPoint;
}

/** `v` turned by `degrees` (the CSS sense: clockwise on screen). */
function turned(v: MorphPoint, degrees: number): MorphPoint {
	const radians = (degrees * Math.PI) / 180;
	const cos = Math.cos(radians);
	const sin = Math.sin(radians);
	return { x: v.x * cos - v.y * sin, y: v.x * sin + v.y * cos };
}

function centreOf(element: PptxElement): MorphPoint {
	return { x: element.x + element.width / 2, y: element.y + element.height / 2 };
}

/**
 * The translation that, about a box centred on `pivot`, lays the local anchor
 * `own` (scaled by `scale`) onto the slide point `target`. `own` and `target`
 * are both read unrotated; `rotation` turns the offset the way the rendered
 * boxes are turned, which is how a rotated pair's text stays glyph over glyph.
 */
function translationOnto(
	target: MorphPoint,
	targetPivot: MorphPoint,
	own: MorphPoint,
	pivot: MorphPoint,
	sx: number,
	sy: number,
	rotation: number,
): { dx: number; dy: number } {
	const v = turned(
		{
			x: target.x - targetPivot.x - sx * (own.x - pivot.x),
			y: target.y - targetPivot.y - sy * (own.y - pivot.y),
		},
		rotation,
	);
	return { dx: targetPivot.x - pivot.x + v.x, dy: targetPivot.y - pivot.y + v.y };
}

/**
 * A turning half's frame: its pivot (offset `tail` from its centre `own`)
 * laid on the slide point `pivotAt`, at `rotation` and `scale`.
 */
function turnFrame(
	pivotAt: MorphPoint,
	own: MorphPoint,
	tail: MorphPoint,
	sx: number,
	sy: number,
): MorphTextTransform {
	return { dx: pivotAt.x - own.x, dy: pivotAt.y - own.y, sx, sy, tail };
}

/** Where a turning half's pivot stands on the slide, at its own rotation. */
function pivotOnSlide(centre: MorphPoint, pivot: MorphPoint, rotation: number): MorphPoint {
	const v = turned({ x: pivot.x - centre.x, y: pivot.y - centre.y }, rotation);
	return { x: centre.x + v.x, y: centre.y + v.y };
}

function tailOf(centre: MorphPoint, pivot: MorphPoint): MorphPoint {
	return { x: pivot.x - centre.x, y: pivot.y - centre.y };
}

/**
 * The incoming half's STARTING transform: it is painted at its own geometry and
 * has to begin with its text laid over the outgoing text.
 */
export function morphTextIncomingStart(
	alignment: MorphTextAlignment,
	to: PptxElement,
): MorphTextTransform {
	const sx = 1 / alignment.scaleX;
	const sy = 1 / alignment.scaleY;
	const pivot = centreOf(to);
	const { turn, centres } = alignment;
	if (turn && centres) {
		const at = pivotOnSlide(centres.from, turn.pivots.from, turn.from);
		return turnFrame(at, pivot, tailOf(pivot, turn.pivots.to), sx, sy);
	}
	const targetPivot = alignment.rotation ? (centres?.from ?? pivot) : pivot;
	return {
		...translationOnto(
			alignment.from,
			targetPivot,
			alignment.to,
			pivot,
			sx,
			sy,
			alignment.rotation ?? 0,
		),
		sx,
		sy,
	};
}

/**
 * The ghost's ENDING transform: it is painted at the outgoing geometry and has
 * to end with its text laid over the incoming text.
 */
export function morphTextGhostEnd(
	alignment: MorphTextAlignment,
	from: PptxElement,
): MorphTextTransform {
	const pivot = centreOf(from);
	const { turn, centres } = alignment;
	if (turn && centres) {
		const at = pivotOnSlide(centres.to, turn.pivots.to, turn.to);
		return turnFrame(
			at,
			pivot,
			tailOf(pivot, turn.pivots.from),
			alignment.scaleX,
			alignment.scaleY,
		);
	}
	const targetPivot = alignment.rotation ? (centres?.to ?? pivot) : pivot;
	return {
		...translationOnto(
			alignment.to,
			targetPivot,
			alignment.from,
			pivot,
			alignment.scaleX,
			alignment.scaleY,
			alignment.rotation ?? 0,
		),
		sx: alignment.scaleX,
		sy: alignment.scaleY,
	};
}

const STILL: MorphTextTransform = { dx: 0, dy: 0, sx: 1, sy: 1 };

/**
 * The incoming half's ENDING transform: its own geometry. Only a turning
 * pair's differs from no transform at all, since its frames carry a pivot.
 */
export function morphTextIncomingEnd(
	alignment: MorphTextAlignment,
	to: PptxElement,
): MorphTextTransform {
	const { turn } = alignment;
	if (!turn || !alignment.centres) {
		return STILL;
	}
	const pivot = centreOf(to);
	return turnFrame(
		pivotOnSlide(pivot, turn.pivots.to, turn.to),
		pivot,
		tailOf(pivot, turn.pivots.to),
		1,
		1,
	);
}

/** The ghost's STARTING transform: its own geometry (see {@link morphTextIncomingEnd}). */
export function morphTextGhostStart(
	alignment: MorphTextAlignment,
	from: PptxElement,
): MorphTextTransform {
	const { turn } = alignment;
	if (!turn || !alignment.centres) {
		return STILL;
	}
	const pivot = centreOf(from);
	return turnFrame(
		pivotOnSlide(pivot, turn.pivots.from, turn.from),
		pivot,
		tailOf(pivot, turn.pivots.from),
		1,
		1,
	);
}

/**
 * The `translate(-tail)` that closes a turning frame's function list, after
 * the element's flips (so the tail is negated along a flipped axis), or `''`.
 */
export function morphTextTail(
	transform: MorphTextTransform,
	element: Pick<PptxElement, 'flipHorizontal' | 'flipVertical'>,
): string {
	const { tail } = transform;
	if (!tail) {
		return '';
	}
	const x = -(element.flipHorizontal ? -tail.x : tail.x);
	const y = -(element.flipVertical ? -tail.y : tail.y);
	return ` translate(${round(x)}px, ${round(y)}px)`;
}

function round(value: number): number {
	return Math.round(value * 1000) / 1000;
}
