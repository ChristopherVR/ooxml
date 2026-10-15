import type { PptxElement, PptxTableData } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { buildRibbonGallery, applyRibbonGalleryItem } from './gallery-registry';
import type { RibbonGalleryContext, RibbonGalleryId } from './gallery-types';

function table(): PptxElement {
	const cell = (text: string) => ({ text });
	const data: PptxTableData = {
		columnWidths: [0.5, 0.3, 0.2],
		rows: [
			{ height: 20, cells: [cell('a'), cell('b'), cell('c')] },
			{ height: 40, cells: [cell('d'), cell('e'), cell('f')] },
			{ height: 60, cells: [cell('g'), cell('h'), cell('i')] },
		],
	};
	return {
		id: 't1',
		type: 'table',
		x: 0,
		y: 0,
		width: 300,
		height: 120,
		tableData: data,
	} as unknown as PptxElement;
}

const run = (id: RibbonGalleryId, ctx: RibbonGalleryContext, item = 'run') =>
	applyRibbonGalleryItem(id, item, ctx);

function dataOf(result: ReturnType<typeof run>): PptxTableData {
	if (result?.kind !== 'element') {
		throw new Error('expected an element result');
	}
	return (result.patch as { tableData: PptxTableData }).tableData;
}

describe('Table Layout commands', () => {
	it('inserts a row above and below the selection, and a column left and right', () => {
		const ctx = { element: table(), selectedTableCells: [{ row: 1, column: 1 }] };
		expect(dataOf(run('tableInsertAbove', ctx)).rows).toHaveLength(4);
		expect(dataOf(run('tableInsertAbove', ctx)).rows[1]?.cells[0]?.text).toBe('');
		expect(dataOf(run('tableInsertBelow', ctx)).rows[2]?.cells[0]?.text).toBe('');
		expect(dataOf(run('tableInsertLeft', ctx)).columnWidths).toHaveLength(4);
		expect(dataOf(run('tableInsertRight', ctx)).rows[0]?.cells).toHaveLength(4);
	});

	it('falls back to the table edge without a selection', () => {
		const ctx = { element: table() };
		expect(dataOf(run('tableInsertAbove', ctx)).rows[0]?.cells[0]?.text).toBe('');
		expect(dataOf(run('tableInsertBelow', ctx)).rows[3]?.cells[0]?.text).toBe('');
	});

	it('deletes the selected rows and columns, and needs a selection', () => {
		const ctx = {
			element: table(),
			selectedTableCells: [
				{ row: 0, column: 0 },
				{ row: 1, column: 0 },
			],
		};
		expect(dataOf(run('tableDeleteRows', ctx)).rows.map((r) => r.cells[0]?.text)).toStrictEqual([
			'g',
		]);
		expect(dataOf(run('tableDeleteColumns', ctx)).columnWidths).toHaveLength(2);
		expect(run('tableDeleteRows', { element: table() })).toBeNull();
		expect(buildRibbonGallery('tableDeleteRows', { element: table() }).disabled).toBe(true);
		expect(buildRibbonGallery('tableDeleteRows', ctx).disabled).toBe(false);
	});

	it('merges a rectangle and splits it again', () => {
		const element = table();
		const cells = [
			{ row: 0, column: 0 },
			{ row: 0, column: 1 },
		];
		expect(buildRibbonGallery('tableMergeCells', { element }).disabled).toBe(true);
		expect(
			buildRibbonGallery('tableMergeCells', { element, selectedTableCells: cells }).disabled,
		).toBe(false);
		const merged = dataOf(run('tableMergeCells', { element, selectedTableCells: cells }));
		expect(merged.rows[0]?.cells[0]?.gridSpan).toBe(2);
		const mergedElement = { ...element, tableData: merged } as PptxElement;
		const anchor = [{ row: 0, column: 0 }];
		expect(
			buildRibbonGallery('tableSplitCells', { element: mergedElement, selectedTableCells: anchor })
				.disabled,
		).toBe(false);
		const split = dataOf(
			run('tableSplitCells', { element: mergedElement, selectedTableCells: anchor }),
		);
		expect(split.rows[0]?.cells[0]?.gridSpan ?? 1).toBe(1);
		expect(buildRibbonGallery('tableSplitCells', { element }).disabled).toBe(true);
	});

	it('distributes rows and columns evenly', () => {
		const ctx = { element: table() };
		expect(dataOf(run('tableDistributeRows', ctx)).rows.map((r) => r.height)).toStrictEqual([
			40, 40, 40,
		]);
		const widths = dataOf(run('tableDistributeColumns', ctx)).columnWidths;
		expect(widths.every((w) => Math.abs(w - 1 / 3) < 1e-9)).toBe(true);
	});

	it('aligns the selected cells, or every cell when the frame is selected', () => {
		const one = dataOf(
			run('tableAlignCenter', { element: table(), selectedTableCells: [{ row: 1, column: 2 }] }),
		);
		expect(one.rows[1]?.cells[2]?.style?.align).toBe('center');
		expect(one.rows[0]?.cells[0]?.style?.align).toBeUndefined();
		const all = dataOf(run('tableAlignBottom', { element: table() }));
		expect(all.rows.flatMap((r) => r.cells).every((c) => c.style?.vAlign === 'bottom')).toBe(true);
	});

	it('sets and clears the text direction', () => {
		const ctx = { element: table(), selectedTableCells: [{ row: 0, column: 0 }] };
		const rotated = dataOf(run('tableTextDirection', ctx, 'rotate270'));
		expect(rotated.rows[0]?.cells[0]?.style?.textDirection).toBe('vert270');
		const next = {
			element: { ...table(), tableData: rotated } as PptxElement,
			selectedTableCells: ctx.selectedTableCells,
		};
		expect(
			buildRibbonGallery('tableTextDirection', next).sections[0]?.items.find((i) => i.applied)?.id,
		).toBe('rotate270');
		const cleared = dataOf(run('tableTextDirection', next, 'horizontal'));
		expect(cleared.rows[0]?.cells[0]?.style?.textDirection).toBeUndefined();
	});

	it('keeps the commands the viewer cannot back disabled with a reason', () => {
		for (const id of [
			'tableSelect',
			'tableViewGridlines',
			'tableCellMargins',
			'tableSize',
			'tableArrange',
		] as const) {
			const gallery = buildRibbonGallery(id, { element: table() });
			expect(gallery.disabled, id).toBe(true);
			expect(gallery.command?.hintKey, id).toBeTruthy();
			expect(run(id, { element: table() }), id).toBeNull();
		}
	});

	it('is disabled for a non-table selection', () => {
		expect(buildRibbonGallery('tableInsertAbove', { element: null }).disabled).toBe(true);
		expect(run('tableInsertAbove', { element: null })).toBeNull();
	});
});
