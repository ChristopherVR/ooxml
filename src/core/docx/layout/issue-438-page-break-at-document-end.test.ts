import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import { loadDocx } from '../index';
import { createFakeMeasurer, layoutDocumentModel } from './index';

async function pageCount(name: string): Promise<number> {
	const { model } = await loadDocx(
		await readFile(new URL(`./fixtures/word-layout-reports/${name}.docx`, import.meta.url)),
	);
	return layoutDocumentModel(model, createFakeMeasurer()).pages.length;
}

// Word 16: a page break in the last paragraph of the document, with nothing after it, makes no
// page (1 page); an empty paragraph after the break is placed on the next page (2 pages).
it('adds no page for a break that ends the document (#438)', async () => {
	expect(await pageCount('issue-438-page-break-at-document-end')).toBe(1);
	expect(await pageCount('issue-438-page-break-after-text-at-document-end')).toBe(1);
});

it('adds the page when a paragraph follows the break (#438 control)', async () => {
	expect(await pageCount('issue-438-page-break-at-document-end-control')).toBe(2);
});
