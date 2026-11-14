import type { PptxTableData } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { mapTableCellEditorKey, resolveTableCellNavigation } from './table-cell-navigation';
import { tableCellsTextFormat, tableCellsTextStylePatch } from './table-cells-format';
import { isTabularPaste, parseTabularText, planTableGridPaste } from './table-tsv-paste';

function grid(rows: number, cols: number): PptxTableData {
	return {
		columnWidths: Array.from({ length: cols }, () => 1 / cols),
		rows: Array.from({ length: rows }, () => ({
			cells: Array.from({ length: cols }, () => ({ text: '' })),
		})),
	};
}

const caret = {
	collapsed: true,
	atStart: false,
	atEnd: false,
	onFirstLine: false,
	onLastLine: false,
};
const key = (name: string, extra: Partial<Record<'shiftKey' | 'ctrlKey', boolean>> = {}) => ({
	key: name,
	shiftKey: false,
	ctrlKey: false,
	metaKey: false,
	altKey: false,
	...extra,
});

describe('resolveTableCellNavigation', () => {
	it('moves Tab to the next cell, wrapping rows, and selects its text', () => {
		const data = grid(2, 2);
		expect(resolveTableCellNavigation(data, { row: 0, col: 1 }, 'next')).toEqual({
			kind: 'move',
			row: 1,
			col: 0,
			caret: 'all',
		});
		expect(resolveTableCellNavigation(data, { row: 1, col: 0 }, 'previous')).toMatchObject({
			row: 0,
			col: 1,
		});
	});

	it('appends a row on Tab in the last cell and stays on Shift+Tab in the first', () => {
		const data = grid(2, 3);
		expect(resolveTableCellNavigation(data, { row: 1, col: 2 }, 'next')).toEqual({
			kind: 'append-row',
			row: 2,
			col: 0,
			caret: 'all',
		});
		expect(resolveTableCellNavigation(data, { row: 0, col: 0 }, 'previous')).toEqual({
			kind: 'stay',
		});
	});

	it('skips merged continuation cells and lands on their owner', () => {
		const data = grid(3, 3);
		data.rows[0]!.cells[0] = { text: 'merged', gridSpan: 2, rowSpan: 2 };
		data.rows[0]!.cells[1] = { text: '', hMerge: true };
		data.rows[1]!.cells[0] = { text: '', vMerge: true };
		data.rows[1]!.cells[1] = { text: '', hMerge: true, vMerge: true };
		expect(resolveTableCellNavigation(data, { row: 0, col: 0 }, 'next')).toMatchObject({
			row: 0,
			col: 2,
		});
		expect(resolveTableCellNavigation(data, { row: 0, col: 0 }, 'down')).toMatchObject({
			row: 2,
			col: 0,
			caret: 'start',
		});
		expect(resolveTableCellNavigation(data, { row: 2, col: 1 }, 'up')).toMatchObject({
			row: 0,
			col: 0,
			caret: 'end',
		});
		expect(resolveTableCellNavigation(data, { row: 1, col: 2 }, 'left')).toMatchObject({
			row: 0,
			col: 2,
			caret: 'end',
		});
	});
});

describe('mapTableCellEditorKey', () => {
	it('maps Tab, Shift+Tab, Ctrl+Tab, Enter, Shift+Enter and Escape', () => {
		expect(mapTableCellEditorKey(key('Tab'), caret)).toEqual({ kind: 'navigate', key: 'next' });
		expect(mapTableCellEditorKey(key('Tab', { shiftKey: true }), caret)).toEqual({
			kind: 'navigate',
			key: 'previous',
		});
		expect(mapTableCellEditorKey(key('Tab', { ctrlKey: true }), caret)?.kind).toBe('tab-character');
		expect(mapTableCellEditorKey(key('Enter'), caret)?.kind).toBe('paragraph');
		expect(mapTableCellEditorKey(key('Enter', { shiftKey: true }), caret)?.kind).toBe('line-break');
		expect(mapTableCellEditorKey(key('Escape'), caret)?.kind).toBe('exit');
		expect(mapTableCellEditorKey(key('b', { ctrlKey: true }), caret)).toEqual({
			kind: 'format',
			property: 'bold',
		});
	});

	it('moves on an arrow only at the edge of the text', () => {
		expect(mapTableCellEditorKey(key('ArrowLeft'), caret)).toBeNull();
		expect(mapTableCellEditorKey(key('ArrowLeft'), { ...caret, atStart: true })).toEqual({
			kind: 'navigate',
			key: 'left',
		});
		expect(mapTableCellEditorKey(key('ArrowRight'), { ...caret, atEnd: true })?.kind).toBe(
			'navigate',
		);
		expect(mapTableCellEditorKey(key('ArrowUp'), { ...caret, onFirstLine: true })).toEqual({
			kind: 'navigate',
			key: 'up',
		});
		expect(mapTableCellEditorKey(key('ArrowDown'), { ...caret, onLastLine: true })).toEqual({
			kind: 'navigate',
			key: 'down',
		});
		// A selection is extended or collapsed by the browser, never a move.
		expect(
			mapTableCellEditorKey(key('ArrowLeft'), { ...caret, atStart: true, collapsed: false }),
		).toBeNull();
		// Right to left: Left walks towards the end of the text.
		expect(mapTableCellEditorKey(key('ArrowLeft'), { ...caret, atEnd: true, rtl: true })).toEqual({
			kind: 'navigate',
			key: 'right',
		});
	});
});

describe('tab-separated paste', () => {
	it('parses spreadsheet text: CRLF rows, tabs, quoted values and the trailing break', () => {
		expect(parseTabularText('a\tb\r\nc\td\r\n')).toEqual([
			['a', 'b'],
			['c', 'd'],
		]);
		expect(parseTabularText('"two\nlines"\t"say ""hi"""\n')).toEqual([['two\nlines', 'say "hi"']]);
		expect(parseTabularText('x')).toEqual([['x']]);
	});

	it('tells a grid from multi-line text', () => {
		expect(isTabularPaste('a\tb')).toBe(true);
		expect(isTabularPaste('a\nb\n')).toBe(true);
		expect(isTabularPaste('a\nb')).toBe(false);
		expect(isTabularPaste('one cell\n')).toBe(false);
	});

	it('plans the rows and columns a paste adds and skips merged positions', () => {
		const data = grid(2, 2);
		data.rows[1]!.cells[1] = { text: '', hMerge: true };
		const plan = planTableGridPaste(data, { row: 1, col: 0 }, [
			['a', 'b', 'c'],
			['d', 'e', 'f'],
		]);
		expect(plan.addRows).toBe(1);
		expect(plan.addColumns).toBe(1);
		expect(plan.writes).toEqual([
			{ row: 1, col: 0, text: 'a' },
			{ row: 1, col: 2, text: 'c' },
			{ row: 2, col: 0, text: 'd' },
			{ row: 2, col: 1, text: 'e' },
			{ row: 2, col: 2, text: 'f' },
		]);
	});
});

describe('table cells format', () => {
	it('reads the first selected cell and formats every run of the range', () => {
		const data = grid(1, 2);
		data.rows[0]!.cells[1] = { text: 'ab', textRuns: [{ text: 'a', bold: true }, { text: 'b' }] };
		const table = { id: 't', type: 'table', x: 0, y: 0, width: 1, height: 1, tableData: data };
		const cells = [{ row: 0, col: 1 }];
		expect(tableCellsTextFormat(table as never, cells).bold).toBe(true);
		const patch = tableCellsTextStylePatch(table as never, cells, { italic: true, fontSize: 20 });
		expect(patch.tableData!.rows[0]!.cells[1]!.textRuns).toEqual([
			{ text: 'a', bold: true, italic: true, fontSize: 20 },
			{ text: 'b', italic: true, fontSize: 20 },
		]);
		expect(patch.tableData!.rows[0]!.cells[0]!.style?.italic).toBeUndefined();
	});
});
