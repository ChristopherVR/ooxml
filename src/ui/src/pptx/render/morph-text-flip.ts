/**
 * What a flip does to carried morph text.
 *
 * PowerPoint never mirrors text: a horizontally flipped shape paints its text
 * as the unflipped shape does, and a vertical flip turns it upside down about
 * its text rectangle (`text-flip-compensation`). So the text of a flipped
 * shape is the text of an UNFLIPPED shape turned by 180 degrees per vertical
 * flip, as long as its text rectangle mirrors onto itself (equal left and
 * right insets, and a geometry text rectangle centred across the box).
 *
 * A flip that changes between the slides therefore morphs like a rotation
 * that changes by 180 degrees. Measured on PowerPoint 16 (`CreateVideo`,
 * 60 fps, 720p): a "Lorem ipsum dolor" text box gaining `flipV` while it moves
 * turns its text 0 -> -13 -> -32 -> -53 -> -75 -> -96 -> -115 -> ... -> -180
 * degrees (anti-clockwise, never mirrored or squashed: the ink stays 318 px
 * long and 26 px tall on every frame), about the middle of its text; a filled
 * rounded rectangle gaining `flipV` does the same with its text while its body
 * flips edge-on. That is the half-turn rule `shortestRotationTarget` already
 * applies to a shape turning 0 -> 180.
 *
 * @module render/morph-text-flip
 */
import type { PptxElement } from 'ooxml-core/pptx';
import { hasTextProperties } from 'ooxml-core/pptx';

import { shortestRotationTarget } from './morph-animation';
import { morphCopiesCannotSplitEffects } from './morph-text-body-copies';
import { paintsBody } from './morph-text-frame-read';
import { textRectCentreX } from './text-flip-compensation';

/** Sub-pixel tolerance for a text rectangle centred across its box. */
const CENTRED_EPSILON_PX = 0.5;

/** The turn (degrees) a vertical flip adds to an element's text. */
export function morphTextFlipTurn(element: Pick<PptxElement, 'flipVertical'>): number {
	return element.flipVertical ? 180 : 0;
}

/**
 * The unflipped element whose text paints exactly where `element`'s does, or
 * `undefined` when there is none (an off-centre text rectangle). An unflipped
 * element is returned as it is.
 */
export function morphTextUpright(element: PptxElement): PptxElement | undefined {
	if (!element.flipHorizontal && !element.flipVertical) {
		return element;
	}
	if (Math.abs(textRectCentreX(element) - element.width / 2) > CENTRED_EPSILON_PX) {
		return undefined;
	}
	const rotation = (element.rotation ?? 0) + morphTextFlipTurn(element);
	return {
		...element,
		rotation: ((rotation % 360) + 360) % 360,
		flipHorizontal: false,
		flipVertical: false,
	} as PptxElement;
}

/**
 * Whether a matched SHAPE (one that paints a body) gains or loses a flip
 * while holding the same text, so its text has to be carried apart from its
 * body (`morph-text-body`): the body flips edge-on, the text never squashes.
 * Such a pair is painted through the overlay even when nothing else about it
 * changes, since the split copies stand in for its ghost.
 */
export function morphShapeTextFlipChanges(from: PptxElement, to: PptxElement): boolean {
	return (
		(Boolean(from.flipHorizontal) !== Boolean(to.flipHorizontal) ||
			Boolean(from.flipVertical) !== Boolean(to.flipVertical)) &&
		hasTextProperties(from) &&
		hasTextProperties(to) &&
		Boolean(from.text?.trim()) &&
		from.text === to.text &&
		(paintsBody(from) || paintsBody(to)) &&
		!morphCopiesCannotSplitEffects(from) &&
		!morphCopiesCannotSplitEffects(to) &&
		morphTextUpright(from) !== undefined &&
		morphTextUpright(to) !== undefined
	);
}

/**
 * The element rotations (degrees) the two halves of a carried text pair are
 * stated at on their first and last frames, each half keeping its own flips.
 * Their TEXT turns the short way from one half's text angle to the other's.
 */
export function morphCarriedTextRotations(
	from: PptxElement,
	to: PptxElement,
): { ghost: [number, number]; incoming: [number, number] } {
	const fromTurn = morphTextFlipTurn(from);
	const toTurn = morphTextFlipTurn(to);
	const fromText = (from.rotation ?? 0) + fromTurn;
	const toText = (to.rotation ?? 0) + toTurn;
	return {
		ghost: [from.rotation ?? 0, shortestRotationTarget(fromText, toText) - fromTurn],
		incoming: [shortestRotationTarget(toText, fromText) - toTurn, to.rotation ?? 0],
	};
}
