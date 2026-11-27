// Layout rules found by comparing with the Word reference corpus (word-corpus.test.ts), pinned
// here with small synthetic inputs so a regression names the rule.
import { describe, expect, it } from 'vitest';
import { createDocument, twips } from '../index';
import { layoutDocument, layoutDocumentModel } from './layout';
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
	it.each([11, 12, 14, 15])('uses the justification rule for Word mode %d', (mode) => {
		const base = createDocument();
		const doc = {
			...base,
			page: { ...base.page, width: 89, marginLeft: 0, marginRight: 0 },
			compatibilityMode: mode,
			blocks: [{ type: 'paragraph' as const, id: 'p', runs: words, align: 'justify' as const }],
		};
		const { pages } = layoutDocumentModel(doc, measurer);
		const block = at(at(at(pages, 0).columns, 0).blocks, 0);
		expect(block.kind).toBe('paragraph');
		if (block.kind === 'paragraph') expect(block.lines).toHaveLength(mode === 15 ? 2 : 3);
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

describe('Word line break references', () => {
	const courier: TextMeasurer = {
		widthOf: (value, font) => value.length * font.sizePx * 0.6,
		lineHeightOf: () => 20,
	};
	const wraps = (value: string, widthPt: number, extra: Partial<LayoutParagraph> = {}) =>
		layoutParagraph(
			text('p', value, { runs: [{ text: value, fontSizePt: 12 }], ...extra }),
			(widthPt * 4) / 3,
			courier,
			noop,
		).lines.map((line) => line.fragments.map((fragment) => fragment.text).join(''));

	it.each(['\u2013', '\u2014'])('allows a line to break after %s in a word', (dash) => {
		// Word COM at 118 points, Courier New 12 pt, legacy left alignment.
		expect(wraps(`0006 alpha beta${dash}gamma delta epsilon`, 118)).toEqual([
			`0006 alpha beta${dash}`,
			'gamma delta',
			'epsilon',
		]);
	});

	it('can break between consecutive em dashes', () => {
		expect(wraps('0000 alpha beta\u2014\u2014gamma delta', 118)).toEqual([
			'0000 alpha beta\u2014',
			'\u2014gamma delta',
		]);
	});

	it('fills a line through the formatting boundary of an overlong word', () => {
		// Word COM at 51 pt, with the first three letters regular and the rest bold.
		expect(
			wraps('', 51, {
				runs: [
					{ text: 'abc', fontSizePt: 12 },
					{ text: 'defghijklmnop', bold: true, fontSizePt: 12 },
				],
			}),
		).toEqual(['abcdefg', 'hijklmn', 'op']);
	});

	it('preserves authored indentation after a paragraph or manual line break', () => {
		expect(wraps('    alpha beta gamma', 94)).toEqual(['    alpha', 'beta gamma']);
		expect(wraps('alpha\n    beta gamma', 94)).toEqual(['alpha', '    beta', 'gamma']);
	});

	it('squeezes the internal spaces of a modern justified line enough for the next word', () => {
		// Word at 168 pt fits 25 Courier characters (180 pt natural) with seven spaces.
		expect(wraps('0714 alpha beta     gamma delta epsilon', 168, { align: 'justify' })).toEqual([
			'0714 alpha beta     gamma',
			'delta epsilon',
		]);
	});

	it.each([
		{ word: 'MARK', shrinkPt: 12, fits: true },
		{ word: 'MARK', shrinkPt: 14.4, fits: false },
		{ word: 'ordonnance', shrinkPt: 17.4, fits: true },
		{ word: 'Z', shrinkPt: 11.4, fits: false },
	])(
		'uses the overflowing word to limit the squeeze for $word at $shrinkPt pt',
		({ word, shrinkPt, fits }) => {
			const prefix = '0000 aa aa aa aa aa aa aa aa aa';
			const first = `${prefix} ${word}`;
			// Word COM mode 15 in Courier New. A longer final word permits more total
			// compression, given exactly the same spaces and font on the line.
			expect(wraps(`${first} extra`, first.length * 7.2 - shrinkPt, { align: 'justify' })[0]).toBe(
				fits ? first : prefix,
			);
		},
	);

	it.each([
		{ split: 1, shrinkPt: 12, fits: true },
		{ split: 2, shrinkPt: 12, fits: true },
		{ split: 3, shrinkPt: 12, fits: true },
		{ split: 2, shrinkPt: 14.4, fits: false },
	])(
		'squeezes for a word split across runs at $split as for the whole word ($shrinkPt pt)',
		({ split, shrinkPt, fits }) => {
			// Word 16 fits "tempor" whether it is one run or "temp" + "or" (issue #364).
			const prefix = '0000 aa aa aa aa aa aa aa aa aa';
			const first = `${prefix} MARK`;
			const lines = wraps('', first.length * 7.2 - shrinkPt, {
				align: 'justify',
				runs: [
					{ text: `${prefix} ${'MARK'.slice(0, split)}`, fontSizePt: 12 },
					{ text: `${'MARK'.slice(split)} extra`, fontSizePt: 12 },
				],
			});
			expect(lines[0]).toBe(fits ? first : prefix);
		},
	);
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
