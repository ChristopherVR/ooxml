import type { PptxTableCell } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { cellParagraphBlocks } from './table-cell-paragraphs';

const twoParagraphs = (paragraphs: PptxTableCell['paragraphs']): PptxTableCell => ({
	text: 'a\nb',
	textRuns: [{ text: 'a' }, { text: '', isParagraphBreak: true }, { text: 'b' }],
	paragraphs,
});

describe('cellParagraphBlocks', () => {
	it('leaves a cell without paragraph layout to the run renderer', () => {
		expect(cellParagraphBlocks({ text: 'a', textRuns: [{ text: 'a' }] })).toBeUndefined();
	});

	it('leaves a cell whose paragraphs only set zero margins to the run renderer', () => {
		// Many writers put `marL="0" indent="0"` on every paragraph.
		const cell: PptxTableCell = {
			text: 'a',
			textRuns: [{ text: 'a', fontSize: 11 }],
			paragraphs: [{ paragraphMarginLeft: 0, paragraphIndent: 0 }],
		};
		expect(cellParagraphBlocks(cell)).toBeUndefined();
	});

	it('leaves cells from writers that repeat the defaults on every paragraph to the run renderer', () => {
		// Google Slides.
		const google = {
			align: 'left' as const,
			rtl: false,
			paragraphMarginLeft: 0,
			paragraphIndent: 0,
			paragraphSpacingBefore: 0,
			paragraphSpacingAfter: 0,
		};
		expect(cellParagraphBlocks(twoParagraphs([google, google]))).toBeUndefined();
		// LibreOffice: single line spacing written out.
		expect(cellParagraphBlocks(twoParagraphs([{ lineSpacing: 1 }, { lineSpacing: 1 }]))).toBe(
			undefined,
		);
	});

	it("leaves a cell whose paragraphs only repeat the cell's alignment to the run renderer", () => {
		// Google Slides writes `algn="l"` on every paragraph.
		expect(cellParagraphBlocks(twoParagraphs([{ align: 'left' }, { align: 'left' }]))).toBe(
			undefined,
		);
		expect(
			cellParagraphBlocks({
				...twoParagraphs([{ align: 'center' }, { align: 'center' }]),
				style: { align: 'center' },
			}),
		).toBeUndefined();
	});

	it('gives each paragraph its runs and its own alignment, indent and line spacing', () => {
		const cell: PptxTableCell = {
			text: 'one\ntwo',
			textRuns: [
				{ text: 'one', fontSize: 8 },
				{ text: '', isParagraphBreak: true },
				{ text: 'two', fontSize: 12 },
				{ text: '', isLineBreak: true },
				{ text: 'more', fontSize: 10 },
			],
			paragraphs: [
				{ align: 'right', paragraphMarginLeft: 10, paragraphIndent: -5, lineSpacingExactPt: 9 },
				{ align: 'dist', lineSpacing: 0.9 },
			],
		};
		expect(cellParagraphBlocks(cell)).toStrictEqual([
			{
				// The first paragraph's alignment is the cell's, so its block
				// takes it from the cell.
				css: {
					marginInlineStart: '10px',
					textIndent: '-5px',
					lineHeight: '12px',
					fontSize: '8pt',
				},
				// Exact spacing: the runs must not grow the line (issue #35).
				runs: [{ text: 'one', fontSize: 8, lineHeight: 0 }],
			},
			{
				// Single spacing is 1.2 lines, and 90% of that is 1.08.
				css: {
					textAlign: 'justify',
					textAlignLast: 'justify',
					lineHeight: 0.9 * 1.2,
					fontSize: '10pt',
				},
				runs: [
					{ text: 'two', fontSize: 12 },
					{ text: '', isLineBreak: true },
					{ text: 'more', fontSize: 10 },
				],
			},
		]);
	});

	it('uses logical margins, so a vertical cell indents and spaces along its own lines', () => {
		const cell: PptxTableCell = {
			...twoParagraphs([{ paragraphMarginLeft: 6, paragraphSpacingAfter: 4 }, {}]),
			style: { textDirection: 'vert' },
		};
		const css = cellParagraphBlocks(cell)?.[0].css;
		expect(css).toMatchObject({ marginInlineStart: '6px', marginBlockEnd: '4px' });
		expect(css).not.toHaveProperty('marginLeft');
		expect(css).not.toHaveProperty('marginBottom');
	});

	it('sets the direction of a right-to-left paragraph and aligns it right by default', () => {
		const css = cellParagraphBlocks(twoParagraphs([{ rtl: true }, {}]))?.map((block) => block.css);
		expect(css).toStrictEqual([{ textAlign: 'right', direction: 'rtl', unicodeBidi: 'embed' }, {}]);
	});

	it("lets a right-to-left first paragraph follow the cell's edited alignment", () => {
		const cell: PptxTableCell = {
			...twoParagraphs([{ rtl: true }, { rtl: true }]),
			style: { align: 'center' },
		};
		expect(cellParagraphBlocks(cell)?.map((block) => block.css)).toStrictEqual([
			{ direction: 'rtl', unicodeBidi: 'embed' },
			{ direction: 'rtl', unicodeBidi: 'embed' },
		]);
	});

	it('places space before and after as shape text does', () => {
		const cell: PptxTableCell = {
			text: 'a\nb\nc',
			textRuns: [
				{ text: 'a' },
				{ text: '', isParagraphBreak: true },
				{ text: 'b' },
				{ text: '', isParagraphBreak: true },
				{ text: 'c' },
			],
			paragraphs: [
				{ paragraphSpacingBefore: 4, paragraphSpacingAfter: 6 },
				{ paragraphSpacingBefore: 8, paragraphSpacingAfter: 2 },
				{ paragraphSpacingBefore: 3, paragraphSpacingAfter: 5 },
			],
		};
		// Both land below the paragraph; the first paragraph's space before and
		// the last one's space after are dropped.
		expect(cellParagraphBlocks(cell)?.map((block) => block.css)).toStrictEqual([
			{ marginBlockEnd: '6px' },
			{ marginBlockEnd: '10px' },
			{ marginBlockEnd: '3px' },
		]);
	});

	it("aligns a paragraph that sets no alignment by its direction, not with the cell's first paragraph", () => {
		const cell: PptxTableCell = {
			...twoParagraphs([{ align: 'right' }, {}]),
			style: { align: 'right' },
		};
		expect(cellParagraphBlocks(cell)?.map((block) => block.css.textAlign)).toStrictEqual([
			undefined,
			'start',
		]);
	});

	it("follows the cell's alignment after it is edited, as the run stream does", () => {
		// The editor's alignment buttons change only `style.align`; most
		// cells carry paragraphs, as PowerPoint writes `marL="0" indent="0"`.
		const cell: PptxTableCell = {
			...twoParagraphs([
				{ align: 'left', paragraphMarginLeft: 0, paragraphIndent: 0 },
				{ align: 'left', paragraphMarginLeft: 0, paragraphIndent: 0 },
			]),
			style: { align: 'center' },
		};
		expect(cellParagraphBlocks(cell)).toBeUndefined();

		const spaced: PptxTableCell = {
			...twoParagraphs([{ align: 'left', paragraphSpacingAfter: 6 }, { align: 'center' }]),
			style: { align: 'right' },
		};
		expect(cellParagraphBlocks(spaced)?.map((block) => block.css.textAlign)).toStrictEqual([
			undefined,
			'center',
		]);
	});

	it('keeps an anchorCtr cell centred on the block path', () => {
		const cell: PptxTableCell = {
			...twoParagraphs([{ paragraphSpacingAfter: 6 }, {}]),
			style: { anchorCtr: true },
		};
		const blocks = cellParagraphBlocks(cell);
		expect(blocks).toHaveLength(2);
		for (const block of blocks ?? []) {
			expect(block.css).not.toHaveProperty('textAlign');
		}
	});

	it('sizes the line box from the runs only when every run sets a size', () => {
		const cell: PptxTableCell = {
			text: 'ab',
			textRuns: [{ text: 'a', fontSize: 8 }, { text: 'b' }],
			paragraphs: [{ lineSpacing: 1.5 }],
		};
		const css = cellParagraphBlocks(cell)?.[0].css;
		expect(css).toBeDefined();
		expect(css).not.toHaveProperty('fontSize');
	});

	it('keeps an empty paragraph as a blank line sized from its end properties', () => {
		const cell: PptxTableCell = {
			text: 'a\n\nb',
			textRuns: [
				{ text: 'a' },
				{ text: '', isParagraphBreak: true },
				{ text: '', isParagraphBreak: true },
				{ text: 'b' },
			],
			paragraphs: [{}, { align: 'center', endParaFontSize: 6 }, {}],
		};
		const blocks = cellParagraphBlocks(cell);
		expect(blocks?.map((block) => block.runs)).toStrictEqual([
			[{ text: 'a' }],
			[{ text: '', isLineBreak: true }],
			[{ text: 'b' }],
		]);
		expect(blocks?.[1].css.fontSize).toBe('6pt');
	});

	it('does not apply a bare negative indent, which would push text out of the cell', () => {
		const cell = twoParagraphs([{ paragraphIndent: -12 }, { align: 'center' }]);
		expect(cellParagraphBlocks(cell)?.[0].css).toStrictEqual({});
	});

	it("keeps an edited cell's first paragraph layout over its plain text", () => {
		const cell: PptxTableCell = { text: 'edited', paragraphs: [{ lineSpacingExactPt: 9 }] };
		expect(cellParagraphBlocks(cell)).toStrictEqual([
			{ css: { lineHeight: '12px' }, runs: [{ text: 'edited', lineHeight: 0 }] },
		]);
	});
});
