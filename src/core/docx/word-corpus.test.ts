// Compares the core's Print Layout against facts recorded from desktop Word for the checked-in
// reference corpus (`__fixtures__/word-corpus/`, written by `scripts/docx-word-corpus.ps1`).
//
// Every corpus document uses Courier New, whose advance is 0.6 em, so the headless measurer below
// is exact for the glyphs and the comparison tests line breaking and pagination, not font metrics.
// Tolerances: page count, each paragraph's start page and its line count and line-start character
// offsets must match Word exactly (monospace text leaves no rounding room). Vertical positions are
// not compared (the line-height model is not Word's); the browser suite owns raster comparison.
// A document that does not match is listed in `acceptedGaps` with the reason; the gap is asserted
// (the test fails if it starts to pass) so it must be removed once fixed.
import { existsSync, readFileSync, readdirSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { loadDocx } from './index';
import { layoutDocumentModel } from './layout/layout';
import type { LayoutBlockBox, LayoutParagraphBox } from './layout/result';
import type { TextMeasurer } from './layout/measure';

const dir = new URL('./__fixtures__/word-corpus/', import.meta.url);
const measurer: TextMeasurer = {
	widthOf: (text, font) => Array.from(text).filter((c) => c !== '\t').length * font.sizePx * 0.6,
	lineHeightOf: (font) => font.sizePx * 1.1328,
};

interface WordParagraph {
	index: number;
	text: string;
	inTable: boolean;
	page: number;
	lineCount: number;
	lineStarts: number[];
}
interface Facts {
	name: string;
	word: string;
	pageCount: number;
	paragraphs: WordParagraph[];
}

/**
 * Known differences from Word. A gap names a document and the Word paragraph number whose problem
 * it covers; those problems are removed from the failure list and the gap is asserted: the test
 * fails when the paragraph starts to match, so the entry must then be deleted.
 */
const acceptedGaps: { id: string; name: string; paragraph: number; reason: string }[] = [
	{
		id: 'bold-justified-indented-first-line',
		name: '01-headings-styles',
		paragraph: 5,
		reason:
			'Bold, justified paragraph with a 36 pt left and 18 pt first-line indent: the first line ' +
			'breaks at 56 characters where Word fits 64, and later lines break 4 to 11 characters ' +
			'early. The cause is not isolated (indent handling, justification, or bold width).',
	},
	{
		id: 'wrapped-float-line-count',
		name: '07-images',
		paragraph: 3,
		reason:
			'Square-wrapped floating picture: one more line than Word (13 against 12). The layout ' +
			'reports its float wrap as an approximation (FLOAT_WRAP_NOTE).',
	},
	{
		id: 'heading-first-line-offset',
		name: '07-images',
		paragraph: 1,
		reason: 'The first line of the heading starts at character 1 instead of 0 in this document.',
	},
	{
		id: 'two-column-section-line-count',
		name: '06-sections-columns',
		paragraph: 10,
		reason:
			'Two-column section after a landscape section: 45 lines where Word has 39 (column ' +
			'width or balancing differs from Word for this section).',
	},
	{
		id: 'toc-field-line-start',
		name: '09-fields-toc',
		paragraph: 2,
		reason:
			'Table-of-contents entry paragraph: our first line starts at source offset 3 (after the ' +
			'field code) where Word reports 0, because Word counts the field start in Range offsets.',
	},
];

/** Letters only, so list labels, note marks, object placeholders and field results do not matter. */
const clean = (text: string) => text.replace(/[^A-Za-z]/g, '').slice(0, 8);

function paragraphBoxes(blocks: LayoutBlockBox[], out: LayoutParagraphBox[]): void {
	for (const block of blocks) {
		if (block.kind === 'paragraph') out.push(block);
		else for (const row of block.rows) for (const cell of row.cells) paragraphBoxes(cell, out);
	}
}

function ourParagraphs(pages: ReturnType<typeof layoutDocumentModel>['pages']) {
	const found: { blockId: string; page: number; text: string; lineStarts: number[] }[] = [];
	for (const page of pages)
		for (const column of page.columns) {
			const boxes: LayoutParagraphBox[] = [];
			paragraphBoxes(column.blocks, boxes);
			for (const box of boxes) {
				const text = box.lines
					.flatMap((line) =>
						line.fragments
							.filter((f) => f.sourceStart === undefined || f.sourceStart !== f.sourceEnd)
							.map((f) => f.text),
					)
					.join('');
				const earlier = found.find((p) => p.blockId === box.blockId);
				// A paragraph that continues on a later page or column is one Word paragraph.
				if (earlier) {
					earlier.lineStarts.push(...box.lines.map((line) => line.sourceStart));
					earlier.text += text;
					continue;
				}
				found.push({
					blockId: box.blockId,
					page: page.index + 1,
					text,
					lineStarts: box.lines.map((line) => line.sourceStart),
				});
			}
		}
	return found;
}

async function compare(name: string): Promise<string[]> {
	const facts = JSON.parse(readFileSync(new URL(`${name}.facts.json`, dir), 'utf8')) as Facts;
	const loaded = await loadDocx(new Uint8Array(readFileSync(new URL(`${name}.docx`, dir))));
	const result = layoutDocumentModel(loaded.model, measurer);
	const problems: string[] = [];
	if (result.pages.length !== facts.pageCount)
		problems.push(`page count ${result.pages.length}, Word ${facts.pageCount}`);
	const ours = ourParagraphs(result.pages);
	let cursor = 0;
	for (const word of facts.paragraphs) {
		const prefix = clean(word.text);
		if (!prefix || word.inTable) continue;
		const at = ours.findIndex((p, i) => i >= cursor && clean(p.text).startsWith(prefix));
		if (at < 0) {
			problems.push(`paragraph ${word.index} "${prefix}" not found in layout`);
			continue;
		}
		cursor = at + 1;
		const mine = ours[at]!;
		if (mine.page !== word.page)
			problems.push(`paragraph ${word.index}: page ${mine.page}, Word ${word.page}`);
		if (mine.lineStarts.length !== word.lineCount)
			problems.push(
				`paragraph ${word.index}: ${mine.lineStarts.length} lines, Word ${word.lineCount}`,
			);
		else if (mine.lineStarts.join() !== word.lineStarts.join())
			problems.push(
				`paragraph ${word.index}: line starts [${mine.lineStarts}], Word [${word.lineStarts}]`,
			);
	}
	return problems;
}

const names = existsSync(dir)
	? readdirSync(dir)
			.filter((file) => file.endsWith('.facts.json'))
			.map((file) => file.replace('.facts.json', ''))
			.sort()
	: [];

describe('Word reference corpus', () => {
	it('has the documents the editor claims to support', () => {
		expect(names.length).toBeGreaterThanOrEqual(12);
	});
	it.each(names)('%s: page count and line breaks match Word', async (name) => {
		let problems = await compare(name);
		for (const gap of acceptedGaps.filter((g) => g.name === name)) {
			const covered = problems.filter((p) => p.startsWith(`paragraph ${gap.paragraph}:`));
			expect(covered, `accepted gap now passes (${gap.id}): remove it`).not.toEqual([]);
			problems = problems.filter((p) => !covered.includes(p));
		}
		expect(problems).toEqual([]);
	});
	it('only lists accepted gaps for documents in the corpus', () => {
		for (const gap of acceptedGaps) expect(names, gap.id).toContain(gap.name);
	});
});
