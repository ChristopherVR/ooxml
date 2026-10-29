// Layout rules found by comparing with the Word reference corpus (word-corpus.test.ts), pinned
// here with small synthetic inputs so a regression names the rule.
import { describe, expect, it } from 'vitest';
import { twips } from '../index';
import { layoutDocument } from './layout';
import { layoutParagraph } from './paragraph-layout';
import type { LayoutDocumentInput, LayoutParagraph } from './input';
import type { TextMeasurer } from './measure';
import { at } from './test-helpers';

// 10px per character and 20px lines; 1 twip is 1/15 px, so 300 twips are 20px.
const measurer: TextMeasurer = { widthOf: (text) => text.length * 10, lineHeightOf: () => 20 };
const noop = () => {};
const text = (
	id: string,
	value: string,
	extra: Partial<LayoutParagraph> = {},
): LayoutParagraph => ({
	kind: 'paragraph',
	id,
	runs: [{ text: value }],
	...extra,
});
const flow = (blocks: LayoutParagraph[], widthPx = 600): LayoutDocumentInput => ({
	sections: [
		{
			page: {
				widthPx,
				heightPx: 1000,
				marginTopPx: 0,
				marginRightPx: 0,
				marginBottomPx: 0,
				marginLeftPx: 0,
			},
			blocks,
		},
	],
});
const boxes = (input: LayoutDocumentInput) =>
	at(at(layoutDocument(input, measurer).pages, 0).columns, 0).blocks;

describe('justified text squeezes its spaces', () => {
	// Width 89 px: "aaaa bbbb" is 90 px, one pixel over, within what its one space can give.
	const words = [{ text: 'aaaa bbbb cccc' }];
	it('fits a word that is slightly too wide when the line is justified', () => {
		const result = layoutParagraph(
			{ kind: 'paragraph', id: 'p', runs: words, align: 'justify' },
			89,
			measurer,
			noop,
		);
		expect(result.lines).toHaveLength(2);
		const [first] = result.lines;
		const last = at(first?.fragments, first!.fragments.length - 1);
		// The squeezed line is closed up to the column edge.
		expect(last.xPx + last.widthPx).toBeCloseTo(89, 5);
	});
	it('does not squeeze left-aligned text', () => {
		const result = layoutParagraph({ kind: 'paragraph', id: 'p', runs: words }, 89, measurer, noop);
		expect(result.lines).toHaveLength(3);
	});
	it('keeps a word that is too wide for the squeeze on the next line', () => {
		const result = layoutParagraph(
			{ kind: 'paragraph', id: 'p', runs: words, align: 'justify' },
			85,
			measurer,
			noop,
		);
		expect(result.lines).toHaveLength(3);
	});
});

describe('space between paragraphs', () => {
	it('puts a paragraph below its own space before', () => {
		const [first] = boxes(flow([text('a', 'one', { spacingBeforeTwips: twips(300) })]));
		expect(first).toMatchObject({ yPx: 20 });
	});
	it('leaves the larger of the space after and the space before, not their sum', () => {
		const result = boxes(
			flow([
				text('a', 'one', { spacingAfterTwips: twips(450) }),
				text('b', 'two', { spacingBeforeTwips: twips(300) }),
			]),
		);
		// 20 px line + the 30 px space after: the 20 px space before is already covered.
		expect(at(result, 1)).toMatchObject({ yPx: 50 });
	});
	it('adds only what the space before exceeds the space after by', () => {
		const result = boxes(
			flow([
				text('a', 'one', { spacingAfterTwips: twips(150) }),
				text('b', 'two', { spacingBeforeTwips: twips(450) }),
			]),
		);
		expect(at(result, 1)).toMatchObject({ yPx: 20 + 30 });
	});
	it('drops the space before of the paragraph after a page break', () => {
		const result = layoutDocument(
			flow([
				text('a', 'one', { runs: [{ text: 'one', breakAfter: 'page' }] }),
				text('b', 'two', { spacingBeforeTwips: twips(300) }),
			]),
			measurer,
		);
		expect(at(at(at(result.pages, 1).columns, 0).blocks, 0)).toMatchObject({ yPx: 0 });
	});
	it('drops the space after a paragraph that contextual spacing joins to the next', () => {
		const joined = { styleId: 'List', contextualSpacing: true, spacingAfterTwips: twips(300) };
		const result = boxes(flow([text('a', 'one', joined), text('b', 'two', joined)]));
		expect(at(result, 1)).toMatchObject({ yPx: 20 });
	});
});

describe('text on both sides of a picture', () => {
	const picture = {
		partName: 'word/media/a.png',
		contentType: 'image/png',
		widthPx: 100,
		heightPx: 50,
		relativeFromH: 'margin',
		alignH: 'center' as const,
		relativeFromV: 'margin',
		offsetYPx: 0,
		wrap: 'square',
	};
	const words = Array.from({ length: 60 }, () => 'word').join(' ');
	const firstLine = (extra: { wrapText?: string }) => {
		const [box] = boxes(
			flow([{ ...text('p', words), floats: [{ ...picture, ...extra }] }]),
		) as unknown as { lines: { fragments: { xPx: number; widthPx: number; text: string }[] }[] }[];
		return at(box?.lines, 0).fragments.filter((fragment) => fragment.text.trim());
	};
	it('continues a line after the picture when the wrap allows both sides', () => {
		// The picture spans 250 to 350 px; with its 12 px distances text stops at 238 and resumes at 362.
		const fragments = firstLine({ wrapText: 'bothSides' });
		expect(fragments.some((f) => f.xPx >= 362)).toBe(true);
		expect(fragments.some((f) => f.xPx + f.widthPx > 238 && f.xPx < 362)).toBe(false);
		expect(Math.max(...fragments.map((f) => f.xPx + f.widthPx))).toBeLessThanOrEqual(600);
	});
	it('keeps text on the wider side without bothSides', () => {
		const fragments = firstLine({});
		expect(fragments.some((f) => f.xPx >= 362) && fragments.some((f) => f.xPx < 238)).toBe(false);
	});
});
