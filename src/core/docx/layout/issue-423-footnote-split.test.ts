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

/** Reference line page, and note lines on each page. */
async function notePages(name: string) {
	const result = await layout(name);
	const referencePage = result.pages.findIndex((page) =>
		page.columns.some((column) =>
			column.blocks.some(
				(block) =>
					block.kind === 'paragraph' &&
					block.lines.some((line) =>
						line.fragments.some((fragment) => fragment.text?.includes('notes:')),
					),
			),
		),
	);
	const lines = result.pages.map((page) =>
		(page.footnotes ?? []).reduce(
			(n, note) => n + note.paragraphs.reduce((m, p) => m + p.lines.length, 0),
			0,
		),
	);
	return { referencePage, lines };
}

it('keeps the reference line and continues a long note on the next page (#423)', async () => {
	const { referencePage, lines } = await notePages('issue-423-long-footnote-split');
	expect(referencePage).toBe(0);
	expect(lines[0]).toBeGreaterThanOrEqual(2);
	expect(lines[1]).toBeGreaterThan(0);
	expect(lines.reduce((a, b) => a + b, 0)).toBe(16);
});

it('keeps a short note whole with its reference line (#423 control)', async () => {
	const { referencePage, lines } = await notePages('issue-423-short-footnote-control');
	expect(referencePage).toBe(0);
	expect(lines).toEqual([3]);
});
