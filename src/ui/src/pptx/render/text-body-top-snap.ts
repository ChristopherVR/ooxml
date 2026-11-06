/**
 * Where PowerPoint puts the top of a text body, and how far the renderer's own
 * body top is from it.
 *
 * PowerPoint snaps the top of a top-anchored text body to a whole device pixel
 * before it lays any line out: the first baseline of a one-line box moved up
 * by exactly one pixel when the shape top crossed `.5` of a pixel, and every
 * line after it sits `round(B + k p)` below that snapped top (`B` the first
 * baseline offset, `p` the line pitch, both unrounded; see
 * `paragraph-baseline`). Measured on PowerPoint 16 exports with the shape top
 * swept in 0.05pt steps (Calibri, Arial, Times New Roman, Consolas and Segoe
 * UI, 12 to 72pt, 100% to 200%) and with insets: the snap point does not
 * depend on the line spacing or the face. From 24pt (32px) up, the flip comes
 * `SNAP_OFFSET_PX` earlier, which is where those sizes' glyph bottoms sit
 * below the snapped line.
 *
 * The renderer paints the shape at its model position (a whole pixel) and adds
 * the body inset as padding, which leaves the body top at a fractional pixel
 * that the browser then rounds on its own. The move that puts it where
 * PowerPoint's snapped top is, is the difference.
 *
 * @module render/text-body-top-snap
 */
import type { PptxElement } from 'ooxml-core/pptx';
import { hasTextProperties } from 'ooxml-core/pptx';

import { DEFAULT_BODY_INSET_TB_PX } from '../constants';
import { snapPlacement } from './device-pixel-group';
import { resolveTextBodyRectPadding } from './text-body-rect';

/** The font size (px) from which the snap point moves. */
const LARGE_TEXT_PX = 32;
/** How much earlier (px) the top snaps for text of {@link LARGE_TEXT_PX} and above. */
const SNAP_OFFSET_PX = 0.3;

/** The painted and exact top of an element's text body, in slide pixels. */
export interface BodyTop {
	/** Where the renderer puts the body top (shape top plus padding). */
	painted: number;
	/** Where PowerPoint puts it, before it snaps it to a pixel. */
	exact: number;
}

/**
 * The text body top of `element`, or `undefined` when its first line is not
 * placed from the top: another vertical anchor, rotated or flipped text or
 * shape, vertical writing, or a group that is rotated or flipped.
 */
export function bodyTopOf(element: PptxElement): BodyTop | undefined {
	if (!hasTextProperties(element)) {
		return undefined;
	}
	const style = element.textStyle;
	const anchor = style?.vAlign ?? 'top';
	if (
		anchor !== 'top' ||
		element.rotation ||
		element.flipHorizontal ||
		element.flipVertical ||
		(style?.textDirection && style.textDirection !== 'horizontal')
	) {
		return undefined;
	}
	const placement = snapPlacement(element);
	if (!placement) {
		return undefined;
	}
	const padding =
		resolveTextBodyRectPadding(element).top + (style?.bodyInsetTop ?? DEFAULT_BODY_INSET_TB_PX);
	return { painted: placement.y + padding, exact: placement.exact.y + padding };
}

/**
 * How far (px) to move a body top that is at `top.painted` so it lands on the
 * pixel PowerPoint snaps `top.exact` to, for a first line of `sizePx` text.
 */
export function bodyTopSnapShift(top: BodyTop, sizePx: number): number {
	const offset = sizePx >= LARGE_TEXT_PX ? SNAP_OFFSET_PX : 0;
	return Math.round(top.exact + offset) - top.painted;
}
