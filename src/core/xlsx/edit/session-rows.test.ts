import { describe, expect, it, vi } from 'vitest';
import { putCell } from '../cells';
import { createWorkbook } from '../workbook';
import { createEditSession } from './session';
import type { WorkbookChange } from './types';

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

describe('a row height refit and calculation', () => {
	/** A1 grows its row when it gets a large font; B1 reads it and C1 subtotals visible rows. */
	function setup(mode?: 'manual') {
		const wb = createWorkbook();
		const ws = wb.sheets[0]!;
		putCell(ws, 0, 0, { value: 1 });
		putCell(ws, 1, 0, { value: 2 });
		putCell(ws, 0, 1, { value: null, formula: 'A1*10' });
		putCell(ws, 0, 2, { value: null, formula: 'SUBTOTAL(109,A1:A2)' });
		if (mode) wb.calcMode = mode;
		const s = createEditSession(wb);
		s.calculateNow({ full: true });
		const full = vi.spyOn(s.calc, 'recalculateAll');
		const rebuild = vi.spyOn(s.calc, 'invalidate');
		const changes: WorkbookChange[] = [];
		s.onChange((change) => changes.push(change));
		return { wb, ws, s, full, rebuild, changes };
	}
	const A1 = { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } };

	it('recalculates only the dependents of the edit, and views still rebuild', () => {
		const { ws, s, full, rebuild, changes } = setup();
		s.applyStyle(0, [A1], { font: { size: 18 } });
		s.setCellInput(0, 0, 0, '5');
		expect(ws.rowInfo.get(0)?.height).toBe(23.25);
		expect(ws.rows.get(0)?.get(1)?.value).toBe(50);
		expect(ws.rows.get(0)?.get(2)?.value).toBe(7);
		expect(full).not.toHaveBeenCalled();
		expect(rebuild).not.toHaveBeenCalled();
		expect(changes[0]?.structural).toBe(true);
	});

	it('undoes and redoes the edit and its height without a full recalculation', () => {
		const { ws, s, full, changes } = setup();
		s.applyStyle(0, [A1], { font: { size: 18 } });
		s.undo();
		expect(ws.rowInfo.get(0)).toBeUndefined();
		s.redo();
		expect(ws.rowInfo.get(0)?.height).toBe(23.25);
		expect(full).not.toHaveBeenCalled();
		expect(changes.map((change) => change.structural)).toEqual([true, true, true]);
	});

	it('keeps a refit from marking a manual workbook for a full rebuild', () => {
		const { wb, ws, s, full } = setup('manual');
		s.applyStyle(0, [A1], { font: { size: 18 } });
		s.setCellInput(0, 0, 0, '5');
		expect(wb.valuesStale).toBe(true);
		expect(ws.rows.get(0)?.get(1)?.value).toBe(10);
		s.calculateNow();
		expect(ws.rows.get(0)?.get(1)?.value).toBe(50);
		expect(full).not.toHaveBeenCalled();
	});

	it('still recalculates formulas that read hidden rows when rows are hidden', () => {
		const { ws, s, full } = setup();
		s.setHidden(0, 'row', [1], true);
		expect(ws.rows.get(0)?.get(2)?.value).toBe(1);
		expect(full).toHaveBeenCalled();
		s.applyStyle(0, [A1], { font: { size: 18 } });
		s.setCellInput(0, 0, 0, '4');
		expect(ws.rows.get(0)?.get(2)?.value).toBe(4);
		s.undo();
		s.undo();
		s.undo();
		expect(ws.rows.get(0)?.get(2)?.value).toBe(3);
	});

	it('recalculates fully when one batch both hides a row and refits another', () => {
		const { ws, s, full } = setup();
		s.batch('Hide and edit', () => {
			s.setHidden(0, 'row', [1], true);
			s.applyStyle(0, [A1], { font: { size: 18 } });
		});
		expect(ws.rowInfo.get(0)?.height).toBe(23.25);
		expect(ws.rows.get(0)?.get(2)?.value).toBe(1);
		expect(full).toHaveBeenCalled();
	});
});
