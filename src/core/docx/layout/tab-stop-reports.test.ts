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
