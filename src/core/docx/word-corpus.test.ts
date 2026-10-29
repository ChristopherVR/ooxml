// Compares the core's Print Layout against facts recorded from desktop Word for the checked-in
// reference corpus (`__fixtures__/word-corpus/`, written by `scripts/docx-word-corpus.ps1`).
//
// Every corpus document uses Courier New, whose advance is 0.6 em, so the headless measurer below
// is exact for the glyphs and the comparison tests line breaking and pagination, not font metrics.
// Tolerances: page count, each paragraph's start page, its line count and its line-start character
// offsets must match Word exactly (monospace text leaves no rounding room). The page position of
// each paragraph's first character (x and y) and of each floating picture's frame must be within
// ONE point: Word reports positions in twips and the layout works in CSS pixels, so a rounded
// pixel (0.75 pt) or line height can differ without being a layout difference, while a real
// difference (a spacing rule, a line height) is several points. The browser suite owns raster
// comparison. A difference that is real is listed in `acceptedGaps` with the reason; the gap is
// asserted (the test fails if it starts to pass) so it must be removed once fixed.
import { existsSync, readFileSync, readdirSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { loadDocx } from './index';
import { layoutDocumentModel } from './layout/layout';
import type { LayoutBlockBox, LayoutPageBox, LayoutParagraphBox } from './layout/result';
import type { TextMeasurer } from './layout/measure';
import { acceptedGaps } from './word-corpus-gaps';

const dir = new URL('./__fixtures__/word-corpus/', import.meta.url);
const POSITION_TOLERANCE_PT = 1;
const PT_PER_PX = 0.75;
const measurer: TextMeasurer = {
	widthOf: (text, font) => Array.from(text).filter((c) => c !== '\t').length * font.sizePx * 0.6,
	lineHeightOf: (font) => font.sizePx * 1.1328,
};

interface WordParagraph {
	index: number;
	text: string;
	inTable: boolean;
	page: number;
	xPt: number;
	yPt: number;
	lineCount: number;
	lineStarts: number[];
}
interface WordShape {
	page: number;
	leftPt: number;
	topPt: number;
	widthPt: number;
	heightPt: number;
	relH: number;
	relV: number;
}
interface Facts {
	name: string;
	pageCount: number;
	paragraphs: WordParagraph[];
	shapes: WordShape[];
}

/** Letters only, so list labels, note marks, object placeholders and field results do not matter. */
const clean = (text: string) => text.replace(/[^A-Za-z]/g, '').slice(0, 8);
const near = (ours: number, word: number) => Math.abs(ours - word) <= POSITION_TOLERANCE_PT;
const pt = (value: number) => value.toFixed(1);

function paragraphBoxes(blocks: LayoutBlockBox[], out: LayoutParagraphBox[]): void {
	for (const block of blocks) {
		if (block.kind === 'paragraph') out.push(block);
		else for (const row of block.rows) for (const cell of row.cells) paragraphBoxes(cell, out);
	}
}

interface OurParagraph {
	blockId: string;
	page: number;
	text: string;
	lineStarts: number[];
	/** Page position of the first character, in points (absent when the line draws no text). */
	xPt?: number;
	yPt: number;
}

function ourParagraphs(pages: LayoutPageBox[]): OurParagraph[] {
	const found: OurParagraph[] = [];
	for (const page of pages)
		for (const column of page.columns) {
			const boxes: LayoutParagraphBox[] = [];
			paragraphBoxes(column.blocks, boxes);
			for (const box of boxes) {
				const drawn = box.lines.flatMap((line) =>
					line.fragments.filter(
						(f) => f.sourceStart === undefined || f.sourceStart !== f.sourceEnd,
					),
				);
				const text = drawn.map((f) => f.text).join('');
				const earlier = found.find((p) => p.blockId === box.blockId);
				// A paragraph that continues on a later page or column is one Word paragraph.
				if (earlier) {
					earlier.lineStarts.push(...box.lines.map((line) => line.sourceStart));
					earlier.text += text;
					continue;
				}
				const first = box.lines[0]?.fragments.find((f) => f.sourceStart !== undefined);
				found.push({
					blockId: box.blockId,
					page: page.index + 1,
					text,
					lineStarts: box.lines.map((line) => line.sourceStart),
					yPt: (page.marginTopPx + box.yPx + (box.lines[0]?.yPx ?? 0)) * PT_PER_PX,
					...(first
						? { xPt: (page.marginLeftPx + column.xPx + (box.xPx ?? 0) + first.xPx) * PT_PER_PX }
						: {}),
				});
			}
		}
	return found;
}

function compareParagraphs(facts: Facts, pages: LayoutPageBox[], problems: string[]): void {
	const ours = ourParagraphs(pages);
	let cursor = 0;
	for (const word of facts.paragraphs) {
		const prefix = clean(word.text);
		if (!prefix || word.inTable) continue;
		const at = ours.findIndex((p, i) => i >= cursor && clean(p.text).startsWith(prefix));
		if (at < 0) {
			problems.push(`paragraph ${word.index} [lines]: "${prefix}" not found in layout`);
			continue;
		}
		cursor = at + 1;
		const mine = ours[at]!;
		const lines = (message: string) => problems.push(`paragraph ${word.index} [lines]: ${message}`);
		if (mine.page !== word.page) lines(`page ${mine.page}, Word ${word.page}`);
		if (mine.lineStarts.length !== word.lineCount)
			lines(`${mine.lineStarts.length} lines, Word ${word.lineCount}`);
		else if (mine.lineStarts.join() !== word.lineStarts.join())
			lines(`line starts [${mine.lineStarts}], Word [${word.lineStarts}]`);
		if (mine.page === word.page && !near(mine.yPt, word.yPt))
			problems.push(`paragraph ${word.index} [y]: ${pt(mine.yPt)} pt, Word ${pt(word.yPt)} pt`);
		if (mine.page === word.page && mine.xPt !== undefined && !near(mine.xPt, word.xPt))
			problems.push(`paragraph ${word.index} [x]: ${pt(mine.xPt)} pt, Word ${pt(word.xPt)} pt`);
	}
}

/** Floating pictures: Word's shapes in document order against the layout's, page by page. */
function compareFloats(facts: Facts, pages: LayoutPageBox[], problems: string[]): void {
	const floats = pages.flatMap((page) => (page.floats ?? []).map((box) => ({ page, box })));
	if (floats.length !== facts.shapes.length)
		problems.push(`shapes [float]: ${floats.length} floats, Word ${facts.shapes.length}`);
	facts.shapes.forEach((shape, index) => {
		const found = floats[index];
		if (!found) return;
		const { page, box } = found;
		const label = `shape ${index + 1} [float]`;
		// The corpus anchors every picture relative to the page, so Word's offsets are page offsets.
		if (shape.relH !== 1 || shape.relV !== 1) problems.push(`${label}: not anchored to the page`);
		const ours = {
			page: page.index + 1,
			leftPt: box.xPx * PT_PER_PX,
			topPt: box.yPx * PT_PER_PX,
			widthPt: box.widthPx * PT_PER_PX,
			heightPt: box.heightPx * PT_PER_PX,
		};
		if (ours.page !== shape.page) problems.push(`${label}: page ${ours.page}, Word ${shape.page}`);
		for (const key of ['leftPt', 'topPt', 'widthPt', 'heightPt'] as const)
			if (!near(ours[key], shape[key]))
				problems.push(`${label}: ${key} ${pt(ours[key])}, Word ${pt(shape[key])}`);
	});
}

async function compare(name: string): Promise<string[]> {
	const facts = JSON.parse(readFileSync(new URL(`${name}.facts.json`, dir), 'utf8')) as Facts;
	const loaded = await loadDocx(new Uint8Array(readFileSync(new URL(`${name}.docx`, dir))));
	const result = layoutDocumentModel(loaded.model, measurer);
	const problems: string[] = [];
	if (result.pages.length !== facts.pageCount)
		problems.push(`page count ${result.pages.length}, Word ${facts.pageCount}`);
	compareParagraphs(facts, result.pages, problems);
	compareFloats(facts, result.pages, problems);
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
	it.each(names)('%s: lines, positions and floats match Word', async (name) => {
		let problems = await compare(name);
		if (process.env.CORPUS_DEBUG) console.log(`${name}\n${problems.join('\n')}`);
		for (const gap of acceptedGaps.filter((g) => g.name === name)) {
			const covered = problems.filter((p) =>
				gap.paragraphs.some((n) => p.startsWith(`paragraph ${n} [${gap.aspect}]:`)),
			);
			// Every paragraph a gap names must still differ, or the gap is stale.
			for (const n of gap.paragraphs)
				expect(
					covered.some((p) => p.startsWith(`paragraph ${n} [${gap.aspect}]:`)),
					`accepted gap now passes for paragraph ${n} (${gap.id}): remove it`,
				).toBe(true);
			problems = problems.filter((p) => !covered.includes(p));
		}
		expect(problems).toEqual([]);
	});
	it('only lists accepted gaps for documents in the corpus', () => {
		for (const gap of acceptedGaps) expect(names, gap.id).toContain(gap.name);
	});
});
