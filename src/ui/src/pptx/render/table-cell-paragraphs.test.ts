import type { PptxTableCell, PptxTableCellParagraph } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import type { FontBoxMeasure } from './paragraph-baseline';
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
					textJustify: 'inter-character',
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
		// PowerPoint adds a paragraph's space after to the next one's space
		// before; the sum sits on one block so normal flow cannot collapse it.
		// The first paragraph's space before and the last one's space after are
		// dropped.
		expect(cellParagraphBlocks(cell)?.map((block) => block.css)).toStrictEqual([
			{ marginBlockEnd: '14px' },
			{ marginBlockEnd: '5px' },
			{},
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

	describe('baseline placement', () => {
		// Calibri's Windows ascent and descent (em), as the browser measures them.
		const calibri: FontBoxMeasure = (font) => {
			const size = Number(/(\d+(?:\.\d+)?)px/u.exec(font)?.[1]);
			return { ascent: 0.952 * size, descent: 0.269 * size };
		};
		const cellCss = { fontFamily: 'Calibri', fontSize: '12pt', lineHeight: 1.2 };
		const cellAt = (sizePt: number, paragraph: PptxTableCellParagraph): PptxTableCell => ({
			text: 'H',
			textRuns: [{ text: 'H', fontSize: sizePt, fontFamily: 'Calibri' }],
			paragraphs: [paragraph],
		});
		const cssOf = (cell: PptxTableCell) =>
			cellParagraphBlocks(cell, { cellCss, measure: calibri })?.[0]?.css;

		it('moves a 150% paragraph down onto 75% of its pitch (PowerPoint 72px, browser 66px)', () => {
			// 40pt is 53.3px, so the pitch is 96px; the browser centres the leading
			// and puts the baseline at 66px, PowerPoint at 72px (measured, table cell).
			expect(cssOf(cellAt(40, { lineSpacing: 1.5 }))).toMatchObject({
				position: 'relative',
				top: '6px',
			});
		});

		it('moves an exact 72pt paragraph of 12pt text down 19px (PowerPoint 72px, browser 53px)', () => {
			expect(cssOf(cellAt(12, { lineSpacingExactPt: 72 }))).toMatchObject({
				position: 'relative',
				top: '19px',
			});
		});

		it('ends the cell at the last line baseline plus descent, as PowerPoint grows its rows', () => {
			// Row content heights measured on PowerPoint 16 (Calibri 24pt, one line):
			// 150% 38.735pt (pitch 43.2pt), exact 24pt 24.335pt, 50% 17.135pt (pitch 14.4pt).
			const margin = (paragraph: PptxTableCellParagraph) =>
				Number.parseFloat(String(cssOf(cellAt(24, paragraph))?.marginBlockEnd));
			expect(margin({ lineSpacing: 1.5 })).toBeCloseTo((38.735 - 43.2) * (96 / 72), 1);
			expect(margin({ lineSpacingExactPt: 24 })).toBeCloseTo((24.335 - 24) * (96 / 72), 1);
			expect(margin({ lineSpacing: 0.5 })).toBeCloseTo((17.135 - 14.4) * (96 / 72), 1);
			// Single spacing already ends where the line pitch does.
			expect(cssOf(cellAt(24, {}))?.marginBlockEnd).toBeUndefined();
		});

		it('adds the extent only to the last paragraph of the cell', () => {
			const cell: PptxTableCell = {
				text: 'H\nH',
				textRuns: [
					{ text: 'H', fontSize: 24, fontFamily: 'Calibri' },
					{ text: '', isParagraphBreak: true },
					{ text: 'H', fontSize: 24, fontFamily: 'Calibri' },
				],
				paragraphs: [{ lineSpacing: 1.5 }, { lineSpacing: 1.5 }],
			};
			const blocks = cellParagraphBlocks(cell, { cellCss, measure: calibri }) ?? [];
			expect(blocks[0]?.css.marginBlockEnd).toBeUndefined();
			expect(Number.parseFloat(String(blocks[1]?.css.marginBlockEnd))).toBeLessThan(-5);
		});

		it('places a larger opening run at an exact pitch level with a line of one size', () => {
			// Issue #35: a 9pt bullet opens a 7pt line at 11pt exact. The runs sit
			// at `line-height: 0`, so the strut alone sets the CSS baseline and
			// PowerPoint's (75% of the pitch) is the same in both cells.
			const lineOf = (bulletPt: number): PptxTableCell => ({
				text: '• text',
				textRuns: [
					{ text: '•', fontSize: bulletPt, fontFamily: 'Calibri' },
					{ text: ' text', fontSize: 7, fontFamily: 'Calibri' },
				],
				paragraphs: [{ lineSpacingExactPt: 11 }],
			});
			const mixed = cellParagraphBlocks(lineOf(9), { cellCss, measure: calibri })?.[0];
			const single = cellParagraphBlocks(lineOf(7), { cellCss, measure: calibri })?.[0];
			expect(single?.css.top).toBeDefined();
			expect(mixed?.css.top).toBe(single?.css.top);
			expect(mixed?.runs.every((run) => run.lineHeight === 0)).toBe(true);
		});

		it('places a paragraph of a cell that spaces nothing, as shape text does', () => {
			expect(cssOf(cellAt(40, {}))?.position).toBe('relative');
		});

		it('leaves the paragraph alone when no font size is known', () => {
			const cell: PptxTableCell = {
				text: 'H',
				textRuns: [{ text: 'H' }],
				paragraphs: [{ lineSpacing: 1.5 }],
			};
			const bare = { cellCss: { fontFamily: 'Calibri' }, measure: calibri };
			expect(cellParagraphBlocks(cell, bare)?.[0]?.css.top).toBeUndefined();
			// The table's default size supplies it.
			expect(cellParagraphBlocks(cell, { ...bare, defaultFontPt: 40 })?.[0]?.css.top).toBe('6px');
		});
	});
});
