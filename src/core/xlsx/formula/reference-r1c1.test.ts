import { describe, expect, it } from 'vitest';
import { book, engine, get, set, E } from './test-helpers';
import { createEditSession } from '../edit/index';
import { createWorkbook } from '../workbook';
import { saveXlsx } from '../write/index';
import { loadXlsx } from '../read/index';
import { getCell } from '../cells';
import { createCalcEngine } from './engine';

// Recorded independently with Excel 16.0 build 20430 at Z1, 2026-10-07.
describe('INDIRECT R1C1 references', () => {
	it.each([
		['COLUMNS(INDIRECT("R1",FALSE))', 16384],
		['ROWS(INDIRECT("C1",FALSE))', 1048576],
		['ROWS(INDIRECT("R1:R3",FALSE))', 3],
		['COLUMNS(INDIRECT("C1:C3",FALSE))', 3],
		['ROWS(INDIRECT("R",FALSE))', 1],
		['COLUMNS(INDIRECT("C",FALSE))', 1],
		['ROWS(INDIRECT("R[+1]C[+1]",FALSE))', 1],
		['ROWS(INDIRECT("R[1]:R[3]",FALSE))', 3],
		['COLUMNS(INDIRECT("C[-2]:C[-1]",FALSE))', 2],
		['ROWS(INDIRECT("R1C1:R3",FALSE))', E.REF],
		['ROWS(INDIRECT("R1:C3",FALSE))', E.REF],
		['ROWS(INDIRECT("R0",FALSE))', E.REF],
		['ROWS(INDIRECT("R1048577",FALSE))', E.REF],
		['COLUMNS(INDIRECT("C16385",FALSE))', E.REF],
		['ROWS(INDIRECT("R0001",FALSE))', 1],
		['ROW(INDIRECT("R[-1]",FALSE))', 1048576],
		['ROW(INDIRECT("R[-1]C1",FALSE))', 1048576],
		['COLUMN(INDIRECT("C[-26]",FALSE))', 16384],
		['ROW(INDIRECT("R[-1048576]C1",FALSE))', E.REF],
		['ROW(INDIRECT("R[-1048577]C1",FALSE))', E.REF],
		['ROW(INDIRECT("R[1048575]C1",FALSE))', 1048576],
		['ROW(INDIRECT("R[1048576]C1",FALSE))', E.REF],
		['COLUMN(INDIRECT("C[16383]",FALSE))', 25],
		['COLUMN(INDIRECT("C[16384]",FALSE))', E.REF],
	] as const)('%s matches recorded Excel behavior', (formula, expected) => {
		const wb = book();
		set(wb, 'Z1', `=${formula}`);
		engine(wb).recalculateAll();
		expect(get(wb, 'Z1')).toEqual(expected);
	});

	it('uses the existing sheet resolver and sparse reference aggregation', () => {
		const wb = book({ 'Sheet 2!A1': 3, 'Sheet 2!A2': 4 }, ['Sheet1', 'Sheet 2']);
		set(wb, 'Z1', '=SUM(INDIRECT("\'Sheet 2\'!C1",FALSE))');
		engine(wb).recalculateAll();
		expect(get(wb, 'Z1')).toBe(7);
	});

	it('preserves a wrapped reference formula and cached value across save/reload', async () => {
		const wb = createWorkbook();
		const formula = 'ROW(INDIRECT("R[-1]C1",FALSE))';
		createEditSession(wb).setCellInput(0, 0, 25, `=${formula}`);
		expect(getCell(wb.sheets[0]!, 0, 25)?.value).toBe(1048576);
		const loaded = await loadXlsx(await saveXlsx(wb));
		expect(getCell(loaded.sheets[0]!, 0, 25)).toMatchObject({ formula, value: 1048576 });
		createCalcEngine(loaded).recalculateAll();
		expect(getCell(loaded.sheets[0]!, 0, 25)?.value).toBe(1048576);
	});
});
