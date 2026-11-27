/**
 * The small building blocks the keynote's slides share: text boxes without insets (so type sits
 * exactly on the grid), cards, pills and the head of a content slide.
 */
import type {
	TextSegmentInput,
	TextStyleInput,
} from '../../../../src/core/pptx/core/builders/sdk/index';
import type { PptxElement } from '../../../../src/core/pptx/core/types/elements';
import type { TextStyle } from '../../../../src/core/pptx/core/types/text';
import type { Box } from '../kit/geometry';
import type { Paragraph, Slide, TextBoxOptions } from '../kit/slide';
import { C, cardFill, softShadow } from './style';

const FLUSH = { left: 0, top: 0, right: 0, bottom: 0 } as const;

/** A text box whose paragraphs share one style; runs may override it. */
export function tx(
	slide: Slide,
	name: string,
	content: string | Paragraph[],
	box: Box,
	style: Partial<TextStyleInput>,
	extra: TextBoxOptions = {},
): PptxElement {
	const paragraphs = typeof content === 'string' ? [content] : content;
	const styled = paragraphs.map((paragraph) =>
		typeof paragraph === 'string'
			? [{ text: paragraph, style }]
			: paragraph.map((run) => ({ text: run.text, style: { ...style, ...run.style } })),
	);
	return slide.text(name, styled, { ...box, insets: FLUSH, autoFit: 'none', ...extra });
}

/** Runs for a shape's text, one array per paragraph, joined with paragraph breaks. */
export function paragraphs(...items: TextSegmentInput[][]): TextSegmentInput[] {
	return items.flatMap((runs, index) =>
		index === 0
			? runs
			: [{ text: '\n', ...(runs[0]?.style ? { style: runs[0].style } : {}) }, ...runs],
	);
}

/** Restyles the runs of a placeholder whose text is `text` (the SDK takes placeholder text plain). */
export function restyle(element: PptxElement, text: string, style: Partial<TextStyle>): void {
	if (!('textSegments' in element)) return;
	for (const segment of element.textSegments ?? [])
		if (segment.text === text) segment.style = { ...segment.style, ...style };
}

/** A panel: the card gradient, a hairline edge and the shared soft shadow. */
export function card(slide: Slide, name: string, box: Box, radius = 5200): PptxElement {
	return slide.shape(name, 'roundRect', {
		...box,
		fill: cardFill,
		stroke: { color: C.line, width: 1 },
		shadow: softShadow,
		adjustments: { adj: radius },
	});
}

/** The kicker, the title and the slide number of a content slide. */
export function head(slide: Slide, kicker: string, title: string | Paragraph[]): void {
	slide.placeholder('body', kicker, 10);
	slide.placeholder('title', title);
	slide.placeholder('sldNum', String(slide.number), 12);
}
