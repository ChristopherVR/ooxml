import type { LayoutFloat } from './input';
import type { LayoutColumnBox, LayoutFloatBox, LayoutLine, LayoutPageBox } from './result';

export const FLOAT_WRAP_NOTE =
	'Text wraps around floating pictures on their largest side, or on both sides of a square, tight or through picture whose wrapText is bothSides (all use the picture’s rectangle, not its outline); lines continuing onto a later page keep their wrapped widths.';

/** What placing a float needs: its size and its horizontal and vertical frame. */
export type FloatFrame = Pick<
	LayoutFloat,
	| 'widthPx'
	| 'heightPx'
	| 'relativeFromH'
	| 'alignH'
	| 'offsetXPx'
	| 'relativeFromV'
	| 'alignV'
	| 'offsetYPx'
>;

export interface Span {
	start: number;
	size: number;
}

/**
 * The areas a float's `wp:positionH`/`wp:positionV` frames name, in one coordinate space: the page
 * for a body paragraph, the cell for an object laid out in a table cell (`cell-floats.ts`).
 */
export interface FloatFrames {
	page: Span;
	margin: Span;
	leftMargin: Span;
	rightMargin: Span;
	column: Span;
	/** Where the character before the anchor starts (`character`); the column's start otherwise. */
	characterPx?: number;
	pageV: Span;
	marginV: Span;
	topMargin: Span;
	bottomMargin: Span;
	paragraph: Span;
	/** The anchor's line (`line`); the paragraph otherwise. */
	line?: Span;
}

function horizontalSpan(float: FloatFrame, frames: FloatFrames): Span {
	switch (float.relativeFromH) {
		case 'page':
			return frames.page;
		case 'margin':
			return frames.margin;
		case 'leftMargin':
		case 'insideMargin':
			return frames.leftMargin;
		case 'rightMargin':
		case 'outsideMargin':
			return frames.rightMargin;
		case 'character':
			return frames.characterPx === undefined
				? frames.column
				: { start: frames.characterPx, size: 0 };
		default:
			// `column`, Word's default.
			return frames.column;
	}
}

function verticalSpan(float: FloatFrame, frames: FloatFrames): Span {
	switch (float.relativeFromV) {
		case 'page':
			return frames.pageV;
		case 'margin':
			return frames.marginV;
		case 'topMargin':
		case 'insideMargin':
			return frames.topMargin;
		case 'bottomMargin':
		case 'outsideMargin':
			return frames.bottomMargin;
		case 'line':
			return frames.line ?? frames.paragraph;
		default:
			// `paragraph`, Word's default.
			return frames.paragraph;
	}
}

/**
 * Whether a floating object is vertically positioned relative to its paragraph (or line),
 * rather than the page or margins.
 */
export function isParagraphAnchoredV(float: Pick<FloatFrame, 'relativeFromV'>): boolean {
	const rel = float.relativeFromV;
	return rel === undefined || rel === 'paragraph' || rel === 'line';
}

/**
 * The frames of a body paragraph's float: the page, its margins, the column and the paragraph;
 * `anchor` (from {@link anchorPoint}) is relative to the paragraph's box.
 */
export function pageFrames(
	page: LayoutPageBox,
	column: Pick<LayoutColumnBox, 'xPx' | 'widthPx'>,
	paragraph: { topPx: number; heightPx: number },
	anchor?: { line: Span; characterPx: number },
): FloatFrames {
	const left = page.marginLeftPx;
	return {
		page: { start: 0, size: page.widthPx },
		margin: { start: left, size: page.widthPx - left - page.marginRightPx },
		leftMargin: { start: 0, size: left },
		rightMargin: { start: page.widthPx - page.marginRightPx, size: page.marginRightPx },
		column: { start: left + column.xPx, size: column.widthPx },
		pageV: { start: 0, size: page.heightPx },
		marginV: {
			start: page.marginTopPx,
			size: page.heightPx - page.marginTopPx - page.marginBottomPx,
		},
		topMargin: { start: 0, size: page.marginTopPx },
		bottomMargin: { start: page.heightPx - page.marginBottomPx, size: page.marginBottomPx },
		paragraph: { start: paragraph.topPx, size: paragraph.heightPx },
		...(anchor
			? {
					line: { start: paragraph.topPx + anchor.line.start, size: anchor.line.size },
					characterPx: left + column.xPx + anchor.characterPx,
				}
			: {}),
	};
}

function place(
	area: Span,
	size: number,
	align: string | undefined,
	offset: number | undefined,
): number {
	if (align === 'center') return area.start + (area.size - size) / 2;
	if (align === 'right' || align === 'bottom' || align === 'outside')
		return area.start + area.size - size;
	if (align) return area.start;
	return area.start + (offset ?? 0);
}

/** A float's top-left and the origin its offsets are measured from, in its frames' space. */
export function placeInFrames(
	float: FloatFrame,
	frames: FloatFrames,
): { xPx: number; yPx: number; originXPx: number; originYPx: number } {
	const h = horizontalSpan(float, frames);
	const v = verticalSpan(float, frames);
	return {
		xPx: place(h, float.widthPx, float.alignH, float.offsetXPx),
		yPx: place(v, float.heightPx, float.alignV, float.offsetYPx),
		originXPx: h.start,
		originYPx: v.start,
	};
}

/**
 * A floating picture's top-left on the page (CSS pixels), from its horizontal and vertical frame,
 * alignment and offset. `paragraph` is the anchor paragraph's page-relative top and height.
 */
export function floatPosition(
	float: FloatFrame,
	page: LayoutPageBox,
	column: Pick<LayoutColumnBox, 'xPx' | 'widthPx'>,
	paragraph: { topPx: number; heightPx: number },
): { xPx: number; yPx: number } {
	const { xPx, yPx } = placeInFrames(float, pageFrames(page, column, paragraph));
	return { xPx, yPx };
}

/** Where a float's `posOffset` is measured from: the start of its horizontal and vertical frame. */
export function floatOrigin(
	float: LayoutFloat,
	page: LayoutPageBox,
	column: Pick<LayoutColumnBox, 'xPx' | 'widthPx'>,
	paragraph: { topPx: number; heightPx: number },
): { xPx: number; yPx: number } {
	const { originXPx, originYPx } = placeInFrames(float, pageFrames(page, column, paragraph));
	return { xPx: originXPx, yPx: originYPx };
}

/**
 * Whether a paragraph-anchored float would extend beyond the page's bottom margin
 * if placed at `paragraph`.
 */
export function floatCrossesBottomMargin(
	page: LayoutPageBox,
	column: Pick<LayoutColumnBox, 'xPx' | 'widthPx'>,
	paragraph: { topPx: number; heightPx: number },
	lines: readonly LayoutLine[],
	float: LayoutFloat,
): boolean {
	if (!isParagraphAnchoredV(float)) return false;
	const frames = pageFrames(page, column, paragraph, anchorPoint(lines, float.sourceOffset));
	const { yPx } = placeInFrames(float, frames);
	const bottomMarginTopPx = page.heightPx - page.marginBottomPx;
	return yPx + float.heightPx > bottomMarginTopPx;
}

/**
 * The anchor's line (relative to its paragraph box's top) and where the character before it starts
 * (relative to the box's left): Word measures a `character` frame from that character's left edge.
 */
export function anchorPoint(
	lines: readonly LayoutLine[],
	sourceOffset: number | undefined,
): { line: Span; characterPx: number } | undefined {
	if (sourceOffset === undefined || !lines.length) return undefined;
	const line =
		lines.find((item) => sourceOffset >= item.sourceStart && sourceOffset < item.sourceEnd) ??
		[...lines].reverse().find((item) => item.sourceStart <= sourceOffset) ??
		lines[0]!;
	const before = sourceOffset - 1;
	let characterPx = line.fragments[0]?.xPx ?? 0;
	for (const fragment of line.fragments) {
		const start = fragment.sourceStart ?? 0;
		const end = fragment.sourceEnd ?? start + fragment.text.length;
		if (start > before) break;
		const length = Math.max(1, end - start);
		characterPx = fragment.xPx + (fragment.widthPx * Math.min(length, before - start)) / length;
	}
	return { line: { start: line.yPx, size: line.heightPx }, characterPx };
}

/** The page box of a placed float: its position and what the print view draws. */
export function floatBox(
	blockId: string,
	float: LayoutFloat,
	index: number,
	position: { xPx: number; yPx: number; originXPx: number; originYPx: number },
	inCell = false,
): LayoutFloatBox {
	return {
		blockId,
		index,
		originXPx: position.originXPx,
		originYPx: position.originYPx,
		xPx: position.xPx,
		yPx: position.yPx,
		widthPx: float.widthPx,
		heightPx: float.heightPx,
		partName: float.partName,
		...(float.altText ? { altText: float.altText } : {}),
		...(float.transform ? { transform: float.transform } : {}),
		...(float.relativeHeight !== undefined ? { relativeHeight: float.relativeHeight } : {}),
		...(float.pictureFill ? { pictureFill: float.pictureFill } : {}),
		...(float.shadow ? { shadow: float.shadow } : {}),
		contentType: float.contentType,
		...(float.chart ? { chart: float.chart } : {}),
		...(float.textBox ? { textBox: float.textBox } : {}),
		...(float.textBoxLayout ? { textBoxLayout: float.textBoxLayout } : {}),
		behindText: Boolean(float.behindText),
		...(float.wrap ? { wrap: float.wrap } : {}),
		...(float.wrapText ? { wrapText: float.wrapText } : {}),
		...(inCell ? { inCell: true } : {}),
	};
}
