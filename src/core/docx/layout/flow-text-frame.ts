// Text frames (`w:framePr`, ECMA-376 17.3.1.11): non-drop-cap framed paragraphs are placed out of
// the text flow at their anchored box position. Consecutive paragraphs with matching `framePr`
// share one frame box; body text wraps around the frame on the next layout pass (`wrap.ts`).
import { twipsToPixels, type ParagraphFrame } from '../index';
import { floatPosition } from './floats';
import { stackParagraphs } from './table-layout';
import type { PageCursor } from './page-cursor';
import type { TextMeasurer } from './measure';
import type { LayoutParagraph, LayoutTextFrame } from './input';
import type { LayoutParagraphBox } from './result';

/** Adapts resolved `w:framePr` attributes into engine-side `LayoutTextFrame` geometry. */
export function adaptTextFrame(frame: ParagraphFrame): LayoutTextFrame {
	const widthPx =
		frame.widthTwips !== undefined && frame.widthTwips > 0
			? twipsToPixels(frame.widthTwips)
			: undefined;
	const heightPx =
		frame.heightTwips !== undefined && frame.heightTwips > 0
			? twipsToPixels(frame.heightTwips)
			: undefined;
	const relativeFromH: LayoutTextFrame['relativeFromH'] =
		frame.hAnchor === 'page' || frame.hAnchor === 'margin' ? frame.hAnchor : 'column';
	const inlineY = frame.yAlign === 'inline';
	const relativeFromV: LayoutTextFrame['relativeFromV'] =
		!inlineY && (frame.vAnchor === 'page' || frame.vAnchor === 'margin')
			? frame.vAnchor
			: 'paragraph';
	const alignV = frame.yAlign === 'inline' ? undefined : frame.yAlign;
	const wrap: LayoutTextFrame['wrap'] =
		frame.wrap === 'none' ? 'none' : frame.wrap === 'notBeside' ? 'notBeside' : 'around';
	return {
		...(widthPx !== undefined ? { widthPx } : {}),
		...(heightPx !== undefined ? { heightPx } : {}),
		...(frame.heightRule ? { heightRule: frame.heightRule } : {}),
		relativeFromH,
		...(frame.xAlign ? { alignH: frame.xAlign } : {}),
		offsetXPx: frame.xTwips !== undefined ? twipsToPixels(frame.xTwips) : 0,
		relativeFromV,
		...(alignV ? { alignV } : {}),
		offsetYPx: inlineY || frame.yTwips === undefined ? 0 : twipsToPixels(frame.yTwips),
		hSpacePx: frame.hSpaceTwips !== undefined ? twipsToPixels(frame.hSpaceTwips) : 0,
		vSpacePx: frame.vSpaceTwips !== undefined ? twipsToPixels(frame.vSpaceTwips) : 0,
		wrap,
	};
}

/** Whether two consecutive paragraphs belong to the same Word text frame. */
export function sameTextFrame(
	a: LayoutTextFrame | undefined,
	b: LayoutTextFrame | undefined,
): boolean {
	if (!a || !b) return false;
	return (
		a.widthPx === b.widthPx &&
		a.heightPx === b.heightPx &&
		a.heightRule === b.heightRule &&
		a.relativeFromH === b.relativeFromH &&
		a.alignH === b.alignH &&
		a.offsetXPx === b.offsetXPx &&
		a.relativeFromV === b.relativeFromV &&
		a.alignV === b.alignV &&
		a.offsetYPx === b.offsetYPx &&
		a.hSpacePx === b.hSpacePx &&
		a.vSpacePx === b.vSpacePx &&
		a.wrap === b.wrap
	);
}

function measureAutoWidth(
	paragraphs: readonly LayoutParagraph[],
	maxWidthPx: number,
	measurer: TextMeasurer,
	note: (message: string) => void,
): number {
	const probed = paragraphs.map((paragraph) =>
		paragraph.align && paragraph.align !== 'left'
			? { ...paragraph, align: 'left' as const }
			: paragraph,
	);
	const stacked = stackParagraphs(probed, maxWidthPx, measurer, note);
	let widest = 0;
	for (const box of stacked.boxes)
		for (const line of box.lines)
			for (const fragment of line.fragments)
				widest = Math.max(widest, fragment.xPx + fragment.widthPx);
	return Math.min(maxWidthPx, widest);
}

function resolvedFrameHeight(frame: LayoutTextFrame, contentHeightPx: number): number {
	if (frame.heightPx === undefined) return contentHeightPx;
	if (frame.heightRule === 'exact') return frame.heightPx;
	if (frame.heightRule === 'atLeast') return Math.max(contentHeightPx, frame.heightPx);
	return contentHeightPx;
}

/**
 * Places a group of consecutive framed paragraphs as one positioned text frame on the current
 * page, without advancing the vertical flow cursor.
 */
export function placeTextFrame(
	cursor: PageCursor,
	paragraphs: LayoutParagraph[],
	frame: LayoutTextFrame,
	measurer: TextMeasurer,
	note: (message: string) => void,
): void {
	const widthPx =
		frame.widthPx !== undefined && frame.widthPx > 0
			? frame.widthPx
			: measureAutoWidth(paragraphs, cursor.columnWidthPx, measurer, note);
	const stacked = stackParagraphs(paragraphs, Math.max(1, widthPx), measurer, note);
	const frameHeightPx = resolvedFrameHeight(frame, stacked.heightPx);
	const page = cursor.page;
	const column = cursor.column;
	const position = floatPosition(
		{
			widthPx,
			heightPx: frameHeightPx,
			relativeFromH: frame.relativeFromH,
			...(frame.alignH ? { alignH: frame.alignH } : {}),
			offsetXPx: frame.offsetXPx,
			relativeFromV: frame.relativeFromV,
			...(frame.alignV ? { alignV: frame.alignV } : {}),
			offsetYPx: frame.offsetYPx,
		},
		page,
		column,
		{ topPx: page.marginTopPx + cursor.y, heightPx: frameHeightPx },
	);
	const boxLeftPx = position.xPx - page.marginLeftPx - column.xPx;
	const frameTopInColumnPx = position.yPx - page.marginTopPx;
	stacked.boxes.forEach((box, index) => {
		const placed: LayoutParagraphBox = {
			...box,
			xPx: boxLeftPx,
			widthPx,
			floating: {
				hSpacePx: frame.hSpacePx,
				vSpacePx: frame.vSpacePx,
				wrap: frame.wrap,
				...(index === 0 ? { frameHeightPx } : {}),
			},
		};
		cursor.place(placed, 0, frameTopInColumnPx + box.yPx - cursor.y);
	});
}
