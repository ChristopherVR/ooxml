import { promises as fs } from 'node:fs';
import path from 'node:path';

import { describe, it, expect } from 'vitest';

import { PptxHandler } from '../../core/PptxHandler';
import type { PptxData, PptxTableCell } from '../../core/types';
import type { TablePptxElement } from '../../core/types/elements';
import { requireFixture } from '../require-fixture';

/**
 * `table-paragraph-layout.pptx` is saved by PowerPoint for Mac 16. Its table
 * has one case per row; the second column's cell sets the paragraph layout
 * the row names.
 */
const fixturePath = requireFixture(
	path.resolve(__dirname, '../fixtures/table-paragraph-layout.pptx'),
);

async function loadTable(
	handler: PptxHandler,
	bytes: Uint8Array,
): Promise<{ data: PptxData; table: TablePptxElement }> {
	const data = await handler.load(
		bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer,
	);
	const table = data.slides[0].elements.find(
		(element) => element.type === 'table',
	) as TablePptxElement;
	return { data, table };
}

function alignmentCell(table: TablePptxElement): PptxTableCell {
	const row = table.tableData!.rows.find(
		(entry) => entry.cells[0].text === 'Alignment per paragraph',
	);
	return row!.cells[1];
}

describe('table cell paragraph layout from a PowerPoint deck', () => {
	it("reads each paragraph's own alignment, line spacing, spacing and indent", async () => {
		const { table } = await loadTable(new PptxHandler(), await fs.readFile(fixturePath));
		const cases = Object.fromEntries(
			table.tableData!.rows.slice(1).map((row) => [row.cells[0].text, row.cells[1].paragraphs]),
		);
		const zero = { paragraphMarginLeft: 0, paragraphIndent: 0 };
		expect(cases['Alignment per paragraph']).toStrictEqual([
			{ align: 'right', ...zero },
			{ align: 'center', ...zero },
		]);
		expect(cases['Exact 10pt line spacing at 9pt']).toStrictEqual([
			{ ...zero, lineSpacingExactPt: 10 },
		]);
		expect(cases['Line spacing 0.8 (multiple)']).toStrictEqual([{ ...zero, lineSpacing: 0.8 }]);
		// 12pt is 16px.
		expect(cases['Space after 12pt between paragraphs']).toStrictEqual(
			Array.from({ length: 3 }, () => ({ ...zero, paragraphSpacingAfter: 16 })),
		);
		// 0.25in is 24px.
		expect(cases['Hanging indent 0.25in']).toStrictEqual([
			{ paragraphMarginLeft: 24, paragraphIndent: -24 },
		]);
	});

	it('keeps paragraphs that align differently through an unedited save', async () => {
		const handler = new PptxHandler();
		const { data } = await loadTable(handler, await fs.readFile(fixturePath));
		const saved = await handler.save(data.slides);
		const { table } = await loadTable(new PptxHandler(), saved);
		expect(alignmentCell(table).paragraphs?.map((paragraph) => paragraph.align)).toStrictEqual([
			'right',
			'center',
		]);
	});

	it('saves the alignment of every paragraph in a cell', async () => {
		const handler = new PptxHandler();
		const { data, table } = await loadTable(handler, await fs.readFile(fixturePath));
		const cell = alignmentCell(table);
		cell.style = { ...cell.style, align: 'left' };
		cell.paragraphs = cell.paragraphs!.map((paragraph) => ({ ...paragraph, align: 'left' }));
		const saved = await handler.save(data.slides);
		const reloaded = alignmentCell((await loadTable(new PptxHandler(), saved)).table);
		expect(reloaded.paragraphs?.map((paragraph) => paragraph.align)).toStrictEqual([
			'left',
			'left',
		]);
	});
});
