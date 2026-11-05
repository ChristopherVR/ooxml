/**
 * Where the text of a flipped shape is painted: the transform and pivot every
 * binding's text body layer applies inside the flipped container.
 *
 * Measured on PowerPoint 16 exports (`Slide.Export`, 1280 x 720) of text boxes
 * and filled shapes, left and right aligned, wrapping and `wrap="none"`, at
 * 0, 30 (authored 330 by the flip) and 120 (authored 240) degrees:
 *
 *   - `flipH` never mirrors the text: it reads at the shape's rotation, in the
 *     same place as in the unflipped shape (a left-aligned line stays left).
 *   - `flipV` turns the text upside down: a left-aligned, top-anchored line
 *     ends at the bottom right of its box, reading right to left; several
 *     lines keep their order bottom-up, so the whole block is turned 180
 *     degrees, not each line.
 *   - Both flips: the shape's own transform is a half turn, and the text turns
 *     with it (the same picture as `flipV` alone on a symmetric box).
 *   - On a right triangle (`rtTriangle`, whose text rectangle sits in its
 *     lower left), `flipV` moves the RECTANGLE with the geometry (to the upper
 *     left) and turns the text 180 degrees inside it. So the turn pivots on the
 *     text rectangle's centre, not the shape's.
 *
 * `getTextCompensationTransform` gives the transform; this module adds the
 * pivot. The text layer is the shape's whole box with the text rectangle and
 * insets as padding, so a horizontal mirror about the box centre would carry
 * an off-centre rectangle to the far side; about the rectangle's centre it
 * stays where the flipped geometry put it.
 *
 * @module render/text-flip-compensation
 */
import type { PptxElement } from 'ooxml-core/pptx';
import { hasTextProperties } from 'ooxml-core/pptx';

import { DEFAULT_BODY_INSET_LR_PX } from '../constants';
import { getTextCompensationTransform } from './element-style-transform';
import { resolveTextBodyRectPadding } from './text-body-rect';

/** Sub-pixel difference under which the text rectangle counts as centred. */
const CENTRED_EPSILON_PX = 0.01;

/**
 * The horizontal centre of `element`'s text rectangle (geometry rectangle and
 * body insets), in px from its own left edge.
 */
export function textRectCentreX(element: PptxElement): number {
	const ts = hasTextProperties(element) ? element.textStyle : undefined;
	const rect = resolveTextBodyRectPadding(element);
	const left = rect.left + (ts?.bodyInsetLeft ?? DEFAULT_BODY_INSET_LR_PX);
	const right = rect.right + (ts?.bodyInsetRight ?? DEFAULT_BODY_INSET_LR_PX);
	return (left + element.width - right) / 2;
}

/**
 * The `transform-origin` the compensation turns about: `'center'` when the
 * text rectangle is centred in the box, else its centre in px.
 */
export function getTextCompensationOrigin(element: PptxElement): string {
	const centre = textRectCentreX(element);
	return Math.abs(centre - element.width / 2) <= CENTRED_EPSILON_PX
		? 'center'
		: `${Math.round(centre * 1000) / 1000}px 50%`;
}

/** The text layer's compensation as a CSS pair, or `undefined` for an unflipped shape. */
export interface TextCompensationStyle {
	transform: string;
	transformOrigin: string;
}

/**
 * The compensation transform and its origin, or `undefined` when the text
 * needs none.
 */
export function getTextCompensationStyle(element: PptxElement): TextCompensationStyle | undefined {
	const transform = getTextCompensationTransform(element);
	return transform ? { transform, transformOrigin: getTextCompensationOrigin(element) } : undefined;
}
