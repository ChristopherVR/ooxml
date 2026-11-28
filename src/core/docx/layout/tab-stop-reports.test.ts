import { readFile } from 'node:fs/promises';
import { describe, expect, it } from 'vitest';
import { loadDocx } from '../index';
import { layoutDocumentModel } from './layout';
import type { TextMeasurer } from './measure';

// Half an em per character, so every width is exact and no font file is needed.
const measurer: TextMeasurer = {
	widthOf: (text, font) => text.length * font.sizePx * 0.5,
	lineHeightOf: (font) => font.sizePx * 1.15,
};

async function lines(name: string) {
	const { model } = await loadDocx(
		await readFile(new URL(`./fixtures/word-layout-reports/${name}.docx`, import.meta.url)),
	);
	const { pages } = layoutDocumentModel(model, measurer);
	return pages.flatMap((page) =>
		page.columns.flatMap((column) =>
			column.blocks.flatMap((block) => (block.kind === 'paragraph' ? block.lines : [])),
		),
	);
}
type Line = Awaited<ReturnType<typeof lines>>[number];
const words = (line: Line) =>
	line.fragments
		.filter((f) => f.text.trim())
		.map((f) => ({ text: f.text, x: f.xPx, w: f.widthPx }));
/** Where the last fragment of the line ends. */
const end = (line: Line) => {
	const last = line.fragments.at(-1)!;
	return last.xPx + last.widthPx;
};

describe('tab stops inherited from paragraph styles (#414)', () => {
	it('places the tabs at the stops the style defines, as with direct stops', async () => {
		const [style, derived] = await lines('issue-414-style-tab-stops');
		const [direct] = await lines('issue-414-direct-tab-stops');
		expect(words(style!)).toEqual(words(direct!));
		const [, value] = words(style!);
		// 4536 twips = 302.4 px; the right stop at 9070 twips = 604.67 px.
		expect(value!.x).toBeCloseTo(302.4, 1);
		expect(end(style!)).toBeCloseTo(604.67, 1);
		// A derived style clears 4536, adds 2000 and keeps 9070.
		expect(words(derived!)[1]!.x).toBeCloseTo(133.33, 1);
		expect(end(derived!)).toBeCloseTo(604.67, 1);
	});
});

const texts = (lineList: Line[]) =>
	lineList.map((line) => line.fragments.map((f) => f.text).join(''));

describe('a right tab stop past the right margin (#385)', () => {
	it('ends the text at the margin on one line, as with a stop at the margin', async () => {
		const bug = await lines('issue-385-right-tab-past-margin');
		const control = await lines('issue-385-right-tab-at-margin');
		expect(control).toHaveLength(3);
		expect(bug).toHaveLength(3);
		// 9070 twips of text area = 604.67 px.
		for (const line of [...bug, ...control]) expect(end(line)).toBeCloseTo(604.67, 1);
		expect(texts(bug)).toEqual(texts(control));
	});
});

describe('text after a right tab that ends at the right margin (#404)', () => {
	// Kerning-like rounding: a whole text measures a little under the sum of its pieces.
	const rounded: TextMeasurer = {
		widthOf: (text, font) => text.length * font.sizePx * 0.5 - 0.01 * text.length ** 2,
		lineHeightOf: (font) => font.sizePx * 1.15,
	};
	async function laidOut(name: string) {
		const { model } = await loadDocx(
			await readFile(new URL(`./fixtures/word-layout-reports/${name}.docx`, import.meta.url)),
		);
		return layoutDocumentModel(model, rounded).pages.flatMap((page) =>
			page.columns.flatMap((column) =>
				column.blocks.flatMap((block) => (block.kind === 'paragraph' ? block.lines : [])),
			),
		);
	}

	it('keeps every tabbed and positional-tab paragraph on one line', async () => {
		expect(await laidOut('issue-404-tab-stop-one-twip-inside')).toHaveLength(5);
		// Five paragraphs with a tab stop at 9072 twips and five with a right `w:ptab`.
		expect(await laidOut('issue-404-tab-stop-at-margin')).toHaveLength(10);
	});
});
