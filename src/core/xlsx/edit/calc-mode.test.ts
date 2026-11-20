import { describe, expect, it } from 'vitest';
import { getCell, putCell } from '../cells';
import { cellView } from '../layout/cell-view';
import { loadXlsx } from '../read/load';
import { createWorkbook } from '../workbook';
import { saveXlsx } from '../write/save';
import { calculateOnOpen } from './calc-mode';
import { createEditSession } from './session';

const setup = () => {
	const wb = createWorkbook({ sheets: ['Sheet1', 'Sheet2'] });
	const s = createEditSession(wb);
	s.setCellValue(0, 0, 0, 2);
	s.setCellInput(0, 0, 1, '=A1*10');
	s.setCellInput(1, 0, 0, '=Sheet1!A1+1');
	return { wb, s };
};
const value = (wb: ReturnType<typeof createWorkbook>, sheet: number, row: number, col: number) =>
	getCell(wb.sheets[sheet]!, row, col)?.value;

describe('manual calculation', () => {
	it('defers recalculation until calculateNow', () => {
		const { wb, s } = setup();
		expect(s.autoRecalc()).toBe(true);
		s.setCalcMode('manual');
		expect(wb.calcMode).toBe('manual');
		expect(s.autoRecalc()).toBe(false);
		expect(s.undoLabel()).toBe('Manual calculation');
		s.setCellValue(0, 0, 0, 5);
		expect(value(wb, 0, 0, 1)).toBe(20);
		const kinds: string[] = [];
		s.onChange((c) => kinds.push(c.label));
		s.calculateNow();
		expect(value(wb, 0, 0, 1)).toBe(50);
		expect(value(wb, 1, 0, 0)).toBe(6);
		expect(kinds).toEqual(['Calculate now']);
	});
	it('calculates one sheet and its dependents, and after structural edits everything', () => {
		const { wb, s } = setup();
		s.setAutoRecalc(false);
		s.setCellValue(0, 0, 0, 7);
		s.calculateSheet(0);
		expect(value(wb, 0, 0, 1)).toBe(70);
		s.insertRows(0, 0, 1);
		s.setCellValue(0, 1, 0, 3);
		expect(value(wb, 0, 1, 1)).toBe(70);
		s.calculateSheet(0);
		expect(value(wb, 0, 1, 1)).toBe(30);
		expect(getCell(wb.sheets[1]!, 0, 0)?.formula).toBe('Sheet1!A2+1');
		// The graph followed the inserted row, so Shift+F9 stays on its sheet; F9 does the rest.
		expect(value(wb, 1, 0, 0)).toBe(8);
		s.calculateNow();
		expect(value(wb, 1, 0, 0)).toBe(4);
		// A structural edit the graph cannot follow (a new sheet) recalculates everything.
		s.setCellValue(0, 1, 0, 1);
		s.addSheet('Extra');
		s.calculateSheet(2);
		expect([value(wb, 0, 1, 1), value(wb, 1, 0, 0)]).toEqual([10, 2]);
	});
	it('recalculates when switching back to automatic and undoes the switch', () => {
		const { wb, s } = setup();
		s.setCalcMode('manual');
		s.setCellValue(0, 0, 0, 4);
		s.setAutoRecalc(true);
		expect(wb.calcMode).toBeUndefined();
		expect(value(wb, 0, 0, 1)).toBe(40);
		s.undo();
		expect(wb.calcMode).toBe('manual');
	});
	it('lets recalc: false win over the workbook mode', () => {
		const wb = createWorkbook({ sheets: ['Sheet1'] });
		const s = createEditSession(wb, { recalc: false });
		expect(s.autoRecalc()).toBe(false);
	});
	it('round-trips calcMode', async () => {
		const { wb, s } = setup();
		s.setCalcMode('manual');
		const back = await loadXlsx(await saveXlsx(wb));
		expect(back.calcMode).toBe('manual');
		s.setCalcMode('auto');
		expect((await loadXlsx(await saveXlsx(wb))).calcMode).toBeUndefined();
	});
});

describe('show formulas', () => {
	it('shows formula text left aligned and round-trips the flag', async () => {
		const { wb, s } = setup();
		s.setSheetView(0, { showFormulas: true });
		const view = cellView(wb, 0, 0, 1);
		expect(view.text).toBe('=A1*10');
		expect(view.hAlign).toBe('left');
		expect(view.isNumber).toBe(false);
		expect(cellView(wb, 0, 0, 0).text).toBe('2');
		const back = await loadXlsx(await saveXlsx(wb));
		expect(back.sheets[0]?.view.showFormulas).toBe(true);
		s.setSheetView(0, { showFormulas: false });
		expect(cellView(wb, 0, 0, 1).text).toBe('20');
	});
});

describe('calculateOnOpen', () => {
	it('runs the full calculation a workbook asks for on open, unless it calculates manually', () => {
		const wb = createWorkbook({ sheets: ['S'] });
		const sheet = wb.sheets[0]!;
		putCell(sheet, 0, 0, { value: 2 });
		putCell(sheet, 0, 1, { value: 99, formula: 'A1*3' });
		expect(calculateOnOpen(wb)).toEqual({ ran: false });
		expect(getCell(sheet, 0, 1)?.value).toBe(99);
		wb.fullCalcOnLoad = true;
		wb.calcMode = 'manual';
		expect(calculateOnOpen(wb)).toEqual({ ran: false });
		expect(getCell(sheet, 0, 1)?.value).toBe(99);
		delete wb.calcMode;
		expect(calculateOnOpen(wb)).toEqual({ ran: true });
		expect(getCell(sheet, 0, 1)?.value).toBe(6);
	});

	it('stops at its time limit and leaves the formulas it did not reach as saved', () => {
		const wb = createWorkbook({ sheets: ['S'] });
		const sheet = wb.sheets[0]!;
		for (let row = 0; row < 500; row++) putCell(sheet, row, 0, { value: -1, formula: 'ROW()' });
		wb.fullCalcOnLoad = true;
		expect(calculateOnOpen(wb, { timeLimitMs: 0 })).toEqual({ ran: true, stopped: 'time' });
		const values = [...Array(500).keys()].map((row) => getCell(sheet, row, 0)?.value);
		expect(values).toContain(-1);
		expect(values.filter((v) => v === -1).length).toBeGreaterThan(400);
		expect(calculateOnOpen(wb)).toEqual({ ran: true });
		expect(getCell(sheet, 499, 0)?.value).toBe(500);
	});
});
