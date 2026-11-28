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

it('draws and spaces the between line of a bordered group (#417)', async () => {
	const [page] = await paragraphs('between-mid-page');
	const [first, second] = page!;
	// Word 16: "Next paragraph" is 43.92 pt below "HEADING" (13.44 + 15 + 0.5 + 15).
	const pitch =
		second!.top +
		second!.line.yPx +
		second!.line.baselinePx! -
		(first!.top + first!.line.yPx + first!.line.baselinePx!);
	// The line itself and the 15 + 0.5 + 15 pt of the group: 30.5 pt more than the line height.
	expect((pitch - first!.line.heightPx) / PX_PER_PT).toBeCloseTo(30.5, 1);
});

it('moves the heading whose between space does not fit at the page end (#417)', async () => {
	const pages = await paragraphs('between-at-page-end');
	expect(pages[0]).toHaveLength(48);
	expect(pages[1]?.map((p) => p.text)).toEqual(['HEADING', 'Next paragraph']);
});

it('keeps the heading on page one without a shared border (#417 control)', async () => {
	const pages = await paragraphs('between-not-shared-control');
	expect(pages[0]).toHaveLength(50);
});
