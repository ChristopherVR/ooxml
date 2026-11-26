// Word's Wrap Text choices for a picture or shape, mapped onto the model's `anchored` flag and
// floating `PicturePlacement`. The writer turns a change into `wp:inline` or `wp:anchor` XML.
import type { InlineImage, PicturePlacement } from './inline-model';

/** Word's Wrap Text menu: In Line with Text, the five wrapping styles, Behind and In Front of Text. */
export type DrawingWrap =
	| 'inline'
	| 'square'
	| 'tight'
	| 'through'
	| 'topAndBottom'
	| 'behind'
	| 'front';

/** Every Wrap Text choice, in Word's menu order. */
export const DRAWING_WRAPS: readonly DrawingWrap[] = [
	'inline',
	'square',
	'tight',
	'through',
	'topAndBottom',
	'behind',
	'front',
];

/** Where a picture lands when it starts floating: offsets from the column and the paragraph. */
export interface FloatOrigin {
	offsetXPx?: number;
	offsetYPx?: number;
}

/** The Wrap Text choice that describes a drawing. */
export function drawingWrapOf(
	anchored: boolean | undefined,
	placement: PicturePlacement | undefined,
): DrawingWrap {
	if (!anchored) return 'inline';
	if (!placement || placement.wrap === 'none') return placement?.behindText ? 'behind' : 'front';
	return placement.wrap;
}

/**
 * The floating placement for a wrap choice. An existing placement keeps its position and the
 * sides text may use; a drawing that starts floating sits at `origin` from the column and the
 * paragraph, as Word places a picture converted from In Line with Text.
 */
export function placementForWrap(
	placement: PicturePlacement | undefined,
	wrap: Exclude<DrawingWrap, 'inline'>,
	origin: FloatOrigin = {},
): PicturePlacement {
	const base: PicturePlacement = placement ?? {
		wrap: 'none',
		relativeFrom: 'column',
		offsetXPx: Math.round(origin.offsetXPx ?? 0),
		relativeFromV: 'paragraph',
		offsetYPx: Math.round(origin.offsetYPx ?? 0),
	};
	const { behindText: _behind, wrapText, ...kept } = base;
	if (wrap === 'behind') return { ...kept, wrap: 'none', behindText: true };
	if (wrap === 'front') return { ...kept, wrap: 'none' };
	if (wrap === 'topAndBottom') return { ...kept, wrap };
	return { ...kept, wrap, ...(wrapText ? { wrapText } : {}) };
}

/** A copy of `image` with the Wrap Text choice applied (`anchored` and `placement`). */
export function setDrawingWrap(
	image: InlineImage,
	wrap: DrawingWrap,
	origin?: FloatOrigin,
): InlineImage {
	if (image.vmlFloat && wrap === 'inline')
		throw new Error('Inline conversion is unavailable for legacy VML pictures.');
	const { placement, anchored: _anchored, ...rest } = image;
	if (wrap === 'inline') return rest;
	return {
		...rest,
		anchored: true,
		placement: placementForWrap(image.anchored ? placement : undefined, wrap, origin),
	};
}
