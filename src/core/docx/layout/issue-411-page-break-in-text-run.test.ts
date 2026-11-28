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

it('starts a page where a page break shares a run with text (#411)', async () => {
	const pages = await pageTexts('issue-411-page-break-in-text-run');
	expect(pages).toHaveLength(3);
	expect(pages[1]?.join('|')).toContain('Page two');
	expect(pages[1]?.join('|')).toContain('Still page two');
	expect(pages[2]?.join('|')).toContain('Page three');
});

it('gives three pages when the page breaks have runs of their own (#411 control)', async () => {
	expect(await pageTexts('issue-411-page-break-own-run')).toHaveLength(3);
});
