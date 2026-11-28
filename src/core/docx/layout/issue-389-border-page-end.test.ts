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

/** Text of every paragraph per page, in order. */
async function pageTexts(name: string) {
	const result = await layout(name);
	return result.pages.map((page) =>
		page.columns.flatMap((column) =>
			column.blocks.flatMap((block) =>
				block.kind === 'paragraph'
					? [
							block.lines
								.flatMap((line) => line.fragments.map((fragment) => fragment.text ?? ''))
								.join(''),
						]
					: [],
			),
		),
	);
}

it('moves a bordered paragraph whose bottom border does not fit (#389)', async () => {
	const pages = await pageTexts('issue-389-border-bottom-space-page-end');
	expect(pages[0]).toHaveLength(47);
	expect(pages[1]?.[0]).toBe('Heading');
});

it('keeps the heading without a bottom border on page one (#389 control)', async () => {
	const pages = await pageTexts('issue-389-border-bottom-space-page-end-control');
	expect(pages[0]).toHaveLength(48);
	expect(pages[0]?.at(-1)).toBe('Heading');
});
