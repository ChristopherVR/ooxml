import { describe, expect, it } from 'vitest';
import { getCell } from '../cells.js';
import { cellView } from '../layout/cell-view.js';
import { loadXlsx } from '../read/load.js';
import { createWorkbook } from '../workbook.js';
import { saveXlsx } from '../write/save.js';
import { createEditSession } from './session.js';

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
		expect(value(wb, 1, 0, 0)).toBe(4);
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
