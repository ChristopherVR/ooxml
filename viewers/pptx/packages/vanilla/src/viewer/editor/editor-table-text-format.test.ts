import type { TablePptxElement } from 'pptx-viewer-core';
import { describe, expect, it } from 'vitest';

import type { Store, ViewerState } from '../state';
import { tableTextFormatPatch } from './editor-table-text-format';

const table = {
	id: 't',
	type: 'table',
	x: 0,
	y: 0,
	width: 100,
	height: 50,
	tableData: {
		columnWidths: [0.5, 0.5],
		rows: [{ cells: [{ text: 'a', textRuns: [{ text: 'a', color: '#FF0000' }] }, { text: 'b' }] }],
	},
} as TablePptxElement;

const storeWith = (partial: Partial<ViewerState>) =>
	({
		get: () => ({ selectedTableCell: null, selectedTableCells: [], ...partial }),
	}) as unknown as Store<ViewerState>;

describe('tableTextFormatPatch', () => {
	it('bolds every run of the selected cells only', () => {
		const store = storeWith({ selectedTableCell: { row: 0, column: 0 } });
		const patch = tableTextFormatPatch(store, table, (current) => ({ bold: !current.bold }));
		const cells = (patch as Partial<TablePptxElement>).tableData!.rows[0]!.cells;
		expect(cells[0]!.textRuns).toEqual([{ text: 'a', color: '#FF0000', bold: true }]);
		expect(cells[1]!.style?.bold).toBeUndefined();
	});

	it('formats the whole table when no cell is selected, and skips other elements', () => {
		const patch = tableTextFormatPatch(storeWith({}), table, () => ({ italic: true }));
		const cells = (patch as Partial<TablePptxElement>).tableData!.rows[0]!.cells;
		expect(cells.every((cell) => cell.style?.italic)).toBe(true);
		const shape = { ...table, type: 'shape' } as never;
		expect(tableTextFormatPatch(storeWith({}), shape, () => ({ italic: true }))).toBeUndefined();
	});
});
