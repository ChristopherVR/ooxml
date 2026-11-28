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

it('moves a keepNext heading whose space after does not fit with its next line (#390)', async () => {
	const pages = await pageTexts('issue-390-keepnext-spacing-after');
	expect(pages[0]).toHaveLength(48);
	expect(pages[1]?.[0]).toBe('Heading');
});

it('moves the heading likewise with keepLines (#390 control)', async () => {
	const pages = await pageTexts('issue-390-keepnext-spacing-after-keeplines-control');
	expect(pages[1]?.[0]).toBe('Heading');
});
