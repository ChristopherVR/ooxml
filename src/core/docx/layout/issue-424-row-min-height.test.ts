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

/** Lines of each table fragment on each page. */
async function tableLines(name: string) {
	const result = await layout(name);
	return result.pages.map((page) =>
		page.columns.flatMap((column) =>
			column.blocks.flatMap((block) =>
				block.kind === 'table'
					? [block.rows.flatMap((row) => row.cells.flat()).reduce((n, p) => n + p.lines.length, 0)]
					: [],
			),
		),
	);
}

it('moves a row whose minimum height does not fit to the next page (#424)', async () => {
	expect(await tableLines('issue-424-row-min-height-at-page-end')).toEqual([[], [5]]);
});

it('still splits the same row without a minimum height (#424 control)', async () => {
	expect(await tableLines('issue-424-row-no-min-height-control')).toEqual([[3], [2]]);
});

it('moves a row whose text fits but whose minimum height does not (#424)', async () => {
	expect(await tableLines('issue-424-row-min-height-content-fits')).toEqual([[], [2]]);
});

it('keeps the row on the page when its text and no minimum height fit (#424 control)', async () => {
	expect(await tableLines('issue-424-row-content-fits-control')).toEqual([[2]]);
});
