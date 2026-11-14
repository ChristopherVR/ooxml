import type { TablePptxElement } from 'pptx-viewer-core';
import { describe, expect, it } from 'vitest';

import type { EditorState } from './editor-state.svelte';
import { tableTextFormatOf, tableTextFormatPatch } from './editor-table-text-format';

const table = {
	id: 't',
	type: 'table',
	x: 0,
	y: 0,
	width: 100,
	height: 50,
	tableData: {
		columnWidths: [0.5, 0.5],
		rows: [{ cells: [{ text: 'a', textRuns: [{ text: 'a', italic: true }] }, { text: 'b' }] }],
	},
} as TablePptxElement;

const editorWith = (cells: Array<{ row: number; col: number }>) =>
	({ tableCells: { cellsFor: () => cells } }) as unknown as EditorState;

describe('tableTextFormatPatch', () => {
	it('formats every run of the selected range only', () => {
		const editor = editorWith([{ row: 0, col: 1 }]);
		expect(tableTextFormatOf(editor, table)?.italic).toBe(false);
		const patch = tableTextFormatPatch(editor, table, (now) => ({ underline: !now.underline }));
		const cells = (patch as Partial<TablePptxElement>).tableData!.rows[0]!.cells;
		expect(cells[1]!.style?.underline).toBe(true);
		expect(cells[0]!.textRuns).toEqual([{ text: 'a', italic: true }]);
	});

	it('formats the whole table with no range and leaves other elements alone', () => {
		const patch = tableTextFormatPatch(editorWith([]), table, () => ({ fontSize: 12 }));
		const cells = (patch as Partial<TablePptxElement>).tableData!.rows[0]!.cells;
		expect(cells[0]!.textRuns).toEqual([{ text: 'a', italic: true, fontSize: 12 }]);
		expect(
			tableTextFormatPatch(editorWith([]), { ...table, type: 'shape' } as never, () => ({})),
		).toBeUndefined();
	});
});
