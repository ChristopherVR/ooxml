import type { TablePptxElement } from 'pptx-viewer-core';
import { describe, expect, it } from 'vitest';

import { tableFontRequestPatch, tableFontStyle } from './ribbon-font-table';

const table = {
	id: 't',
	type: 'table',
	x: 0,
	y: 0,
	width: 100,
	height: 50,
	tableData: {
		columnWidths: [0.5, 0.5],
		rows: [{ cells: [{ text: 'a', textRuns: [{ text: 'a', bold: true }] }, { text: 'b' }] }],
	},
} as TablePptxElement;

describe('ribbon Font group on a table', () => {
	it('toggles bold on every run of the Shift+click range from its first cell', () => {
		const selection = {
			elementId: 't',
			rowIndex: 0,
			columnIndex: 0,
			selectedCells: [
				{ row: 0, col: 0 },
				{ row: 0, col: 1 },
			],
		};
		expect(tableFontStyle(table, selection)?.bold).toBe(true);
		const patch = tableFontRequestPatch(table, selection, 'home.font.bold', undefined);
		const cells = (patch as Partial<TablePptxElement>).tableData!.rows[0]!.cells;
		expect(cells[0]!.textRuns).toEqual([{ text: 'a', bold: false }]);
		expect(cells[1]!.style?.bold).toBe(false);
	});

	it('sets the size on the whole table with no cell selected, in points', () => {
		const patch = tableFontRequestPatch(table, null, 'home.font.fontSize', 20);
		const cells = (patch as Partial<TablePptxElement>).tableData!.rows[0]!.cells;
		expect(cells.map((cell) => cell.style?.fontSize)).toEqual([20, 20]);
		expect(
			tableFontRequestPatch({ ...table, type: 'shape' } as never, null, 'home.font.bold', 0),
		).toBeUndefined();
	});

	it('applies highlight, character spacing, text shadow and change case to selected cells', () => {
		const selection = {
			elementId: 't',
			rowIndex: 0,
			columnIndex: 0,
			selectedCells: [{ row: 0, col: 0 }],
		};
		const hl = tableFontRequestPatch(table, selection, 'home.font.highlightColor', '#ffff00');
		expect((hl as Partial<TablePptxElement>).tableData!.rows[0]!.cells[0]!.textRuns).toMatchObject([
			{ text: 'a', highlightColor: '#ffff00' },
		]);

		const spc = tableFontRequestPatch(table, selection, 'home.font.characterSpacing', 75);
		expect((spc as Partial<TablePptxElement>).tableData!.rows[0]!.cells[0]!.textRuns).toMatchObject(
			[{ text: 'a', characterSpacing: 75 }],
		);

		const shOn = tableFontRequestPatch(table, selection, 'home.font.shadow', undefined);
		const withShadow = { ...table, ...shOn } as TablePptxElement;
		expect(tableFontStyle(withShadow, selection)?.textShadowColor).toBe('#000000');
		const shOff = tableFontRequestPatch(withShadow, selection, 'home.font.shadow', undefined);
		expect(
			(shOff as Partial<TablePptxElement>).tableData!.rows[0]!.cells[0]!.style?.textShadowColor,
		).toBeUndefined();

		const cased = tableFontRequestPatch(table, selection, 'home.font.changeCase', 'upper');
		expect((cased as Partial<TablePptxElement>).tableData!.rows[0]!.cells[0]!.text).toBe('A');
		expect((cased as Partial<TablePptxElement>).tableData!.rows[0]!.cells[1]!.text).toBe('b');
	});
});
