import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import { loadDocx } from '../index';
import { createFakeMeasurer, layoutDocumentModel } from './index';

async function layout(name: string) {
	const { model } = await loadDocx(
		await readFile(new URL(`./fixtures/word-layout-reports/${name}.docx`, import.meta.url)),
	);
	return layoutDocumentModel(model, createFakeMeasurer());
}

const PX_PER_PT = 96 / 72;

/** Paragraph boxes with their first line, page by page. */
async function paragraphs(name: string) {
	const result = await layout(name);
	return result.pages.map((page) =>
		page.columns.flatMap((column) =>
			column.blocks.flatMap((block) =>
				block.kind === 'paragraph'
					? [
							{
								text: block.lines[0]?.fragments.map((f) => f.text ?? '').join('') ?? '',
								top: block.yPx,
								line: block.lines[0]!,
							},
						]
					: [],
			),
		),
	);
}

it('puts the baseline of an exact line at 80 % of its height (#419)', async () => {
	const [page] = await paragraphs('exact-line-baseline');
	// Word 16: 15 pt exact line, Arial 10 pt, baseline 12.05 pt below the top of the line.
	const heights = [10, 15, 20, 15, 15, 15];
	page!.forEach(({ line }, index) => {
		expect(line.heightPx / PX_PER_PT).toBeCloseTo(heights[index]!, 2);
		expect(line.baselinePx! / PX_PER_PT).toBeCloseTo(heights[index]! * 0.8, 2);
	});
});
