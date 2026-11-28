import { readFile } from 'node:fs/promises';
import { describe, expect, it } from 'vitest';
import { loadDocx } from '../index';
import { layoutDocumentModel } from './layout';
import type { TextMeasurer } from './measure';

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

describe('the number takes the paragraph mark formatting (#421)', () => {
	it('ignores the first run of the text', async () => {
		const [eightOverTwenty, plain, courier, twentyOverEight] = (
			await lines('issue-421-number-from-first-run')
		).map((line) => line.fragments[0]!);
		// Word: 8 pt for a mark of 8 pt, the 11 pt default without mark formatting.
		expect(eightOverTwenty).toMatchObject({ text: '1.', fontSizePt: 8 });
		expect(plain).toMatchObject({ text: '2.', fontSizePt: 11 });
		expect(plain?.fontFamily).toBe('Arial');
		expect(courier).toMatchObject({ fontFamily: 'Courier New', bold: true });
		expect(courier?.italic).toBeFalsy();
		expect(twentyOverEight).toMatchObject({ text: '4.', fontSizePt: 20 });
	});

	it('keeps the number of a paragraph whose mark and text agree', async () => {
		const [small, large] = (await lines('issue-421-number-mark-control')).map(
			(line) => line.fragments[0]!,
		);
		expect(small?.fontSizePt).toBe(8);
		expect(large?.fontSizePt).toBe(20);
	});
});
