import type { LayoutParagraphBorders } from './input';
import type { LayoutParagraphFrame } from './result';

/**
 * Word reserves a border's width above and below a bordered paragraph whose top and bottom lines
 * cast a shadow: the box sits that far under the previous paragraph, and the shadow strip under
 * it takes the same room before the next one.
 */
export function paragraphShadowInsets(borders: LayoutParagraphBorders | undefined) {
	return {
		topPx: borders?.top?.shadow ? borders.top.widthPx : 0,
		bottomPx: borders?.bottom?.shadow ? borders.bottom.widthPx : 0,
	};
}

/** A filled rectangle, relative to the top-left corner of a paragraph frame. */
export interface ParagraphShadowRect {
	xPx: number;
	yPx: number;
	widthPx: number;
	heightPx: number;
	color: string;
}

/**
 * The shadow of a paragraph frame `heightPx` tall: the box moved right and down by the width of
 * its right and bottom lines, drawn outside the box, as Word draws `w:shadow` on those lines.
 */
export function paragraphShadowRects(
	frame: Pick<LayoutParagraphFrame, 'widthPx' | 'borders'>,
	heightPx: number,
): ParagraphShadowRect[] {
	const { right, bottom } = frame.borders ?? {};
	const rects: ParagraphShadowRect[] = [];
	if (right?.shadow && heightPx > right.widthPx)
		rects.push({
			xPx: frame.widthPx,
			yPx: right.widthPx,
			widthPx: right.widthPx,
			heightPx: heightPx - right.widthPx,
			color: right.color,
		});
	if (bottom?.shadow)
		rects.push({
			xPx: bottom.widthPx,
			yPx: heightPx,
			widthPx: frame.widthPx - bottom.widthPx + (right?.shadow ? right.widthPx : 0),
			heightPx: bottom.widthPx,
			color: bottom.color,
		});
	return rects;
}
