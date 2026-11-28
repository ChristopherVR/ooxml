import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import { loadDocx } from '../index';
import { createFakeMeasurer, layoutDocumentModel } from './index';

async function lineCounts(name: string): Promise<number[]> {
	const { model } = await loadDocx(
		await readFile(new URL(`./fixtures/word-layout-reports/${name}.docx`, import.meta.url)),
	);
	const { pages } = layoutDocumentModel(model, createFakeMeasurer());
	return pages.flatMap((page) =>
		page.columns.flatMap((column) =>
			column.blocks.flatMap((block) => (block.kind === 'paragraph' ? [block.lines.length] : [])),
		),
	);
}

// Word 16 (compatibility mode 15): a text-wrapping break that ends a paragraph, in a run of its own
// or after the text of a run, ends the line and the paragraph mark takes one more, empty, line.
it('gives a paragraph ending in a line break an empty last line (#429)', async () => {
	expect(await lineCounts('issue-429-line-break-at-paragraph-end')).toEqual([2, 1, 2, 1, 1]);
});

it('keeps one line per paragraph without the breaks (#429 control)', async () => {
	expect(await lineCounts('issue-429-line-break-at-paragraph-end-control')).toEqual([
		1, 1, 1, 1, 1,
	]);
});
