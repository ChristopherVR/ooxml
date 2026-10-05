import { describe, expect, it } from 'vitest';
import { putCell } from '../cells.js';
import { createWorkbook } from '../workbook.js';
import { createEditSession } from './session.js';

const all = { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } };

describe('session automatic row height', () => {
	it('grows a row when the font size grows and undoes with the edit', () => {
		const wb = createWorkbook();
		const s = createEditSession(wb, { recalc: false });
		s.setCellInput(0, 0, 0, 'Title');
		expect(wb.sheets[0]?.rowInfo.get(0)).toBeUndefined();
		s.applyStyle(0, [all], { font: { size: 18 } });
		expect(wb.sheets[0]?.rowInfo.get(0)?.height).toBe(23.25);
		expect(s.undoLabel()).not.toBe('Row height');
		s.undo();
		expect(wb.sheets[0]?.rowInfo.get(0)).toBeUndefined();
		s.redo();
		expect(wb.sheets[0]?.rowInfo.get(0)?.height).toBe(23.25);
	});

	it('grows for wrapped text and keeps custom heights', () => {
		const wb = createWorkbook();
		const s = createEditSession(wb, { recalc: false, measureText: (t) => t.length * 7 });
		s.applyStyle(0, [all], { alignment: { wrapText: true } });
		s.setCellInput(0, 0, 0, 'word '.repeat(30).trim());
		expect(wb.sheets[0]?.rowInfo.get(0)?.height ?? 0).toBeGreaterThan(30);
		s.setRowHeight(0, [1], 40);
		s.setCellInput(0, 1, 0, 'x');
		expect(wb.sheets[0]?.rowInfo.get(1)?.height).toBe(40);
	});

	it('can be turned off', () => {
		const wb = createWorkbook();
		const sheet = wb.sheets[0];
		if (!sheet) throw new Error('no sheet');
		putCell(sheet, 0, 0, { value: 'a' });
		const s = createEditSession(wb, { recalc: false, autoRowHeight: false });
		s.applyStyle(0, [all], { font: { size: 24 } });
		expect(sheet.rowInfo.get(0)).toBeUndefined();
	});

	it('wires drawing commands', () => {
		const wb = createWorkbook();
		const sheet = wb.sheets[0];
		if (!sheet) throw new Error('no sheet');
		sheet.drawings.push({
			kind: 'unsupported',
			description: 'shape',
			anchor: { from: { row: 0, col: 0, rowOffset: 0, colOffset: 0 } },
		});
		const s = createEditSession(wb, { recalc: false });
		s.setDrawingAnchor(0, 0, { from: { row: 4, col: 2, rowOffset: 0, colOffset: 0 } });
		expect(sheet.drawings[0]?.anchor.from.row).toBe(4);
		s.deleteDrawing(0, 0);
		expect(sheet.drawings).toHaveLength(0);
		s.undo();
		s.undo();
		expect(sheet.drawings[0]?.anchor.from.row).toBe(0);
	});
});
