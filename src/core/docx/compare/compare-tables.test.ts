import { readFile } from 'node:fs/promises';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { acceptAllRevisions, rejectAllRevisions } from '../revision-commands';
import { bodyText, COMPARE_OPTIONS, docx, p, tbl } from '../test-support/compare-docs';
import { compareDocuments } from './compare';

/** Compares, saves, and checks that rejecting gives `before` and accepting gives `after`. */
async function roundTrip(original: string, revised: string, before: string[], after: string[]) {
	const result = await compareDocuments(await docx(original), await docx(revised), COMPARE_OPTIONS);
	expect(bodyText(rejectAllRevisions(result.model))).toEqual(before);
	expect(bodyText(acceptAllRevisions(result.model))).toEqual(after);
	return result;
}

describe('compareDocuments with tables', () => {
	it('keeps an inserted table above the paragraph that follows it', async () => {
		const result = await roundTrip(
			p('Kept.', 'Old'),
			p('Kept.') + tbl(['T']) + p('New'),
			['Kept.', 'Old'],
			['Kept.', 'T', 'New'],
		);
		expect(result.model.blocks.map((block) => block.type)).toEqual([
			'paragraph',
			'table',
			'paragraph',
		]);
	});

	it('keeps text added above a changed table above it', async () => {
		const result = await roundTrip(
			p('Kept.') + tbl(['A', 'one']) + p(''),
			p('Kept.', 'Added.') + tbl(['A', 'two']) + p(''),
			['Kept.', 'A', 'one', ''],
			['Kept.', 'Added.', 'A', 'two', ''],
		);
		expect(result.model.blocks.map((block) => block.type)).toEqual([
			'paragraph',
			'paragraph',
			'table',
			'paragraph',
		]);
	});

	it('shows a table with a column added as replaced, so it still saves', async () => {
		const result = await roundTrip(
			p('Kept.') + tbl(['A', 'B']) + p(''),
			p('Kept.') + tbl(['A', 'B', 'C']) + p(''),
			['Kept.', 'A', 'B', ''],
			['Kept.', 'A', 'B', 'C', ''],
		);
		expect(result.warnings.join(' ')).toContain('deleted and inserted whole');
	});

	it('shows a paragraph removed or added just above a table', async () => {
		await roundTrip(
			p('Kept.', 'Gone.') + tbl(['T']) + p(''),
			p('Kept.') + tbl(['T']) + p(''),
			['Kept.', 'Gone.', 'T', ''],
			['Kept.', 'T', ''],
		);
		await roundTrip(
			p('Kept.') + tbl(['T']) + p(''),
			p('Kept.', 'Added.') + tbl(['T']) + p(''),
			['Kept.', 'T', ''],
			['Kept.', 'Added.', 'T', ''],
		);
	});

	it('leaves tables with nested tables alone when they did not change', async () => {
		const bytes = await readFile(
			join(__dirname, '..', '__fixtures__', 'nested-three-row', 'nested-three-short.docx'),
		);
		const result = await compareDocuments(bytes, bytes, COMPARE_OPTIONS);
		expect(result.summary).toEqual({ insertions: 0, deletions: 0, formatChanges: 0 });
	});
});
