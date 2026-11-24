import { tableTextShadowUpdate } from 'ooxml-ui/pptx';
import type { TablePptxElement } from 'pptx-viewer-core';
import { describe, expect, it, vi } from 'vitest';

import { createStore } from '../state';
import type { Store, ViewerState } from '../state';
import { createStructuredEditorOperations } from './editor-structured-operations';
import { tableTextChangeCasePatch, tableTextFormatPatch } from './editor-table-text-format';

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

	it('applies highlight, spacing, shadow and changeCase to selected table cells', () => {
		const store = storeWith({ selectedTableCell: { row: 0, column: 0 } });
		const styled = tableTextFormatPatch(store, table, () => ({
			highlightColor: '#ffff00',
			characterSpacing: 150,
			...tableTextShadowUpdate(false),
		}));
		const cell0 = (styled as Partial<TablePptxElement>).tableData!.rows[0]!.cells[0]!;
		expect(cell0.textRuns).toMatchObject([
			{
				text: 'a',
				color: '#FF0000',
				highlightColor: '#ffff00',
				characterSpacing: 150,
				textShadowColor: '#000000',
			},
		]);

		const cased = tableTextChangeCasePatch(store, table, 'upper');
		expect((cased as Partial<TablePptxElement>).tableData!.rows[0]!.cells[0]!.text).toBe('A');
		expect((cased as Partial<TablePptxElement>).tableData!.rows[0]!.cells[1]!.text).toBe('b');
	});

	it('restores selectedTableCell and selectedTableCells on Escape exit', () => {
		const store = createStore({
			slides: [{ id: 's1', slideNumber: 1, elements: [table] }],
			currentSlide: 0,
			selectedElementId: 't',
			selectedElementIds: ['t'],
			selectedTableCell: null,
			selectedTableCells: [],
		} as unknown as ViewerState);
		const ops = createStructuredEditorOperations({
			store,
			pushHistory: vi.fn(),
			commitChange: vi.fn(),
		});
		ops.endTableCellEdit('t', 0, 1, { reason: 'exit', segments: undefined });
		expect(store.get().selectedTableCell).toStrictEqual({ row: 0, column: 1 });
		expect(store.get().selectedTableCells).toStrictEqual([{ row: 0, column: 1 }]);
	});
});
