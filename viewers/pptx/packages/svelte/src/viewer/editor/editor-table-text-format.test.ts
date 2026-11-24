import { tableTextShadowUpdate } from 'ooxml-ui/pptx';
import type { TablePptxElement } from 'pptx-viewer-core';
import { describe, expect, it, vi } from 'vitest';

import { createEditCommits } from '../components/viewer-body-commits';
import type { EditorState } from './editor-state.svelte';
import {
	tableTextChangeCasePatch,
	tableTextFormatOf,
	tableTextFormatPatch,
} from './editor-table-text-format';
import { TableCellSelection } from './table-cell-selection.svelte';

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

	it('applies highlight, spacing, shadow and changeCase to selected table cells', () => {
		const editor = editorWith([{ row: 0, col: 0 }]);
		const styled = tableTextFormatPatch(editor, table, () => ({
			highlightColor: '#ffff00',
			characterSpacing: 75,
			...tableTextShadowUpdate(false),
		}));
		const next = { ...table, ...styled } as TablePptxElement;
		expect(tableTextFormatOf(editor, next)).toMatchObject({
			highlightColor: '#ffff00',
			characterSpacing: 75,
			textShadowColor: '#000000',
		});

		const cased = tableTextChangeCasePatch(editor, table, 'upper');
		expect((cased as Partial<TablePptxElement>).tableData!.rows[0]!.cells[0]!.text).toBe('A');
		expect((cased as Partial<TablePptxElement>).tableData!.rows[0]!.cells[1]!.text).toBe('b');
	});

	it('restores cell selection on Escape exit from table cell editing', () => {
		const tableCells = new TableCellSelection();
		const editor = {
			activeElements: [table],
			tableCells,
			applyElementPatch: vi.fn(),
			transformCommittedText: (text: string) => text,
		} as unknown as EditorState;
		const commits = createEditCommits(editor);
		commits.commitTableCell('t', 0, 1, '', { reason: 'exit', segments: undefined });
		expect(tableCells.cellsFor('t')).toStrictEqual([{ row: 0, col: 1 }]);
	});
});
