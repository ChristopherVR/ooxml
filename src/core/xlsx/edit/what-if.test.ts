import { readFileSync } from 'node:fs';
import path from 'node:path';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { getCell } from '../cells';
import type { Workbook } from '../model';
import { loadXlsx } from '../read/load';
import { createWorkbook } from '../workbook';
import { saveXlsx } from '../write/save';
import { checkDataTable, dataTableFormulaText, DATA_TABLE_PART_ERROR } from './data-table';
import { checkGoalSeek } from './goal-seek';
import { createEditSession } from './session';

const FIXTURE = path.join(import.meta.dirname, '..', '__fixtures__', 'openpyxl-whatif.xlsx');
const value = (wb: Workbook, row: number, col: number, sheet = 0) =>
	getCell(wb.sheets[sheet]!, row, col)?.value;

const book = () => {
	const wb = createWorkbook({ sheets: ['Sheet1'] });
	const s = createEditSession(wb);
	return { wb, s };
};

describe('Goal Seek', () => {
	it('solves a linear model through the recalculation engine in one undo step', () => {
		const { wb, s } = book();
		s.setCellValue(0, 0, 0, 2);
		s.setCellInput(0, 1, 0, '=A1*3+4');
		const r = s.goalSeek({
			sheet: 0,
			target: { row: 1, col: 0 },
			goal: 25,
			changing: { row: 0, col: 0 },
		});
		expect(r.found).toBe(true);
		expect(value(wb, 0, 0) as number).toBeCloseTo(7, 3);
		expect(value(wb, 1, 0) as number).toBeCloseTo(25, 3);
		s.undo();
		expect(value(wb, 0, 0)).toBe(2);
		expect(value(wb, 1, 0)).toBe(10);
	});

	it('solves a loan payment for the amount', () => {
		const { wb, s } = book();
		s.setCellValue(0, 0, 0, 100000);
		s.setCellInput(0, 1, 0, '=-PMT(0.06/12,360,A1)');
		const r = s.goalSeek({
			sheet: 0,
			target: { row: 1, col: 0 },
			goal: 1500,
			changing: { row: 0, col: 0 },
		});
		expect(r.found).toBe(true);
		expect(value(wb, 0, 0) as number).toBeCloseTo(250187, -1);
	});

	it('reports when no solution exists and leaves the best guess', () => {
		const { s } = book();
		s.setCellValue(0, 0, 0, 3);
		s.setCellInput(0, 1, 0, '=A1^2+1');
		const r = s.goalSeek({
			sheet: 0,
			target: { row: 1, col: 0 },
			goal: 0,
			changing: { row: 0, col: 0 },
		});
		expect(r.found).toBe(false);
	});

	it('honours the workbook iteration settings', () => {
		const { wb, s } = book();
		wb.iterate = { count: 1, delta: 0.001 };
		s.setCellValue(0, 0, 0, 1);
		s.setCellInput(0, 1, 0, '=A1^3');
		const r = s.goalSeek({
			sheet: 0,
			target: { row: 1, col: 0 },
			goal: 1000,
			changing: { row: 0, col: 0 },
		});
		expect(r.iterations).toBeLessThanOrEqual(1);
	});

	it('validates the request', () => {
		const { wb, s } = book();
		s.setCellValue(0, 0, 0, 2);
		s.setCellInput(0, 1, 0, '=A1');
		const req = { sheet: 0, target: { row: 1, col: 0 }, goal: 1, changing: { row: 0, col: 0 } };
		expect(checkGoalSeek(wb, req)).toBeNull();
		expect(checkGoalSeek(wb, { ...req, target: { row: 0, col: 0 } })).toBe('same-cell');
		expect(
			checkGoalSeek(wb, { ...req, changing: { row: 1, col: 0 }, target: { row: 0, col: 0 } }),
		).toBe('target-not-formula');
		expect(checkGoalSeek(wb, { ...req, goal: Number.NaN })).toBe('goal-not-number');
		expect(() =>
			s.goalSeek({ ...req, target: { row: 0, col: 0 }, changing: { row: 1, col: 0 } }),
		).toThrow();
	});
});

describe('data tables', () => {
	const setup = () => {
		const { wb, s } = book();
		s.setCellValue(0, 0, 0, 3);
		s.setCellValue(0, 1, 0, 4);
		// Column-oriented table: values A5:A7, formulas B4:C4.
		s.setCellInput(0, 3, 1, '=A1*10');
		s.setCellInput(0, 3, 2, '=A1*A1+A2');
		[1, 2, 3].forEach((v, i) => s.setCellValue(0, 4 + i, 0, v));
		return { wb, s };
	};
	const area = { start: { row: 3, col: 0 }, end: { row: 6, col: 2 } };

	it('computes a one-variable table without changing the real input', () => {
		const { wb, s } = setup();
		s.createDataTable(0, { range: area, colInput: 'A1' });
		expect([4, 5, 6].map((r) => value(wb, r, 1))).toEqual([10, 20, 30]);
		expect([4, 5, 6].map((r) => value(wb, r, 2))).toEqual([5, 8, 13]);
		expect(value(wb, 0, 0)).toBe(3);
		expect(value(wb, 3, 1)).toBe(30);
		expect(wb.sheets[0]!.dataTables?.[0]).toMatchObject({ r1: 'A1' });
		expect(dataTableFormulaText(wb.sheets[0]!, 4, 1)).toBe('{=TABLE(,A1)}');
	});

	it('recomputes when a precedent changes and on undo', () => {
		const { wb, s } = setup();
		s.createDataTable(0, { range: area, colInput: 'A1' });
		s.setCellValue(0, 1, 0, 100);
		expect(value(wb, 4, 2)).toBe(101);
		s.undo();
		expect(value(wb, 4, 2)).toBe(5);
		s.undo();
		expect(wb.sheets[0]!.dataTables).toBeUndefined();
	});

	it('computes a two-variable table', () => {
		const { wb, s } = book();
		s.setCellValue(0, 0, 0, 3);
		s.setCellValue(0, 1, 0, 4);
		s.setCellInput(0, 3, 3, '=A1*A2');
		[1, 2, 3].forEach((v, i) => {
			s.setCellValue(0, 3, 4 + i, v);
			s.setCellValue(0, 4 + i, 3, (i + 1) * 10);
		});
		s.createDataTable(0, {
			range: { start: { row: 3, col: 3 }, end: { row: 6, col: 6 } },
			rowInput: 'A1',
			colInput: '$A$2',
		});
		expect([4, 5, 6].map((r) => [4, 5, 6].map((c) => value(wb, r, c)))).toEqual([
			[10, 20, 30],
			[20, 40, 60],
			[30, 60, 90],
		]);
		expect(wb.sheets[0]!.dataTables?.[0]).toMatchObject({ dt2D: true, r1: 'A1', r2: 'A2' });
		expect(dataTableFormulaText(wb.sheets[0]!, 5, 5)).toBe('{=TABLE(A1,A2)}');
	});

	it('refuses to change part of a table but lets the whole table be cleared', () => {
		const { wb, s } = setup();
		s.createDataTable(0, { range: area, colInput: 'A1' });
		expect(() => s.setCellInput(0, 5, 1, '7')).toThrow(DATA_TABLE_PART_ERROR);
		expect(() =>
			s.clearRange(0, { start: { row: 4, col: 1 }, end: { row: 5, col: 1 } }, 'contents'),
		).toThrow(DATA_TABLE_PART_ERROR);
		s.clearRange(0, { start: { row: 4, col: 1 }, end: { row: 6, col: 2 } }, 'contents');
		expect(wb.sheets[0]!.dataTables).toBeUndefined();
		expect(value(wb, 5, 1) ?? null).toBeNull();
	});

	it('validates the dialog input', () => {
		const { wb } = setup();
		expect(checkDataTable(wb, 0, { range: area })).toBe('no-input');
		expect(checkDataTable(wb, 0, { range: area, rowInput: 'zz' })).toBe('bad-input');
		expect(checkDataTable(wb, 0, { range: area, rowInput: 'B5' })).toBe('input-inside');
		expect(
			checkDataTable(wb, 0, {
				range: { start: { row: 0, col: 0 }, end: { row: 0, col: 3 } },
				rowInput: 'A1',
			}),
		).toBe('range-too-small');
		expect(checkDataTable(wb, 0, { range: area, rowInput: 'Other!A1' })).toBe('bad-input');
	});

	it('reads, recalculates and round-trips the openpyxl fixture', async () => {
		const wb = await loadXlsx(new Uint8Array(readFileSync(FIXTURE)));
		const sheet = wb.sheets[0]!;
		expect(sheet.dataTables).toEqual([
			{ ref: { start: { row: 1, col: 3 }, end: { row: 3, col: 4 } }, r1: 'A1' },
			{
				ref: { start: { row: 1, col: 7 }, end: { row: 3, col: 9 } },
				r1: 'A1',
				r2: 'A2',
				dt2D: true,
				dtr: true,
			},
		]);
		const s = createEditSession(wb);
		s.calculateNow({ full: true });
		s.setCellValue(0, 0, 0, 3);
		expect([1, 2, 3].map((r) => value(wb, r, 3))).toEqual([10, 20, 30]);
		expect([1, 2, 3].map((r) => value(wb, r, 4))).toEqual([1, 4, 9]);
		expect([1, 2, 3].map((r) => [7, 8, 9].map((c) => value(wb, r, c)))).toEqual([
			[10, 20, 30],
			[20, 40, 60],
			[30, 60, 90],
		]);
		const saved = await saveXlsx(wb);
		const xml = await (
			await JSZip.loadAsync(saved)
		)
			.file('xl/worksheets/sheet1.xml')!
			.async('string');
		expect(xml).toContain('t="dataTable"');
		expect(xml).toMatch(/<f t="dataTable" ref="D2:E4" r1="A1"\/>/);
		expect(xml).toMatch(/<f t="dataTable" ref="H2:J4" dt2D="1" dtr="1" r1="A1" r2="A2"\/>/);
		const again = await loadXlsx(saved);
		expect(again.sheets[0]!.dataTables).toEqual(sheet.dataTables);
		expect(value(again, 3, 9)).toBe(90);
		expect(again.warnings.filter((w) => /data table/i.test(w))).toEqual([]);
	});
});

describe('scenarios', () => {
	it('adds, shows, edits and deletes scenarios with undo', () => {
		const { wb, s } = book();
		s.setCellValue(0, 0, 0, 5);
		s.setCellValue(0, 1, 0, 'text');
		s.setCellInput(0, 2, 0, '=A1*2');
		const low = s.addScenario(0, { name: 'Low', cells: 'A1:A2', comment: 'base' });
		expect(wb.sheets[0]!.scenarios?.list[low]?.cells).toEqual([
			{ ref: 'A1', value: '5' },
			{ ref: 'A2', value: 'text' },
		]);
		const high = s.addScenario(0, {
			name: 'High',
			cells: 'A1,A2',
			values: { A1: '50', A2: 'big' },
		});
		expect(() => s.addScenario(0, { name: 'low', cells: 'A1' })).toThrow(/already exists/);
		s.showScenario(0, high);
		expect(value(wb, 0, 0)).toBe(50);
		expect(value(wb, 1, 0)).toBe('big');
		expect(value(wb, 2, 0)).toBe(100);
		expect(wb.sheets[0]!.scenarios).toMatchObject({ show: 1, current: 1 });
		s.undo();
		expect(value(wb, 0, 0)).toBe(5);
		s.editScenario(0, high, { name: 'Higher', cells: 'A1', values: { A1: '99' } });
		expect(wb.sheets[0]!.scenarios?.list[1]).toMatchObject({
			name: 'Higher',
			cells: [{ ref: 'A1', value: '99' }],
		});
		s.deleteScenario(0, 0);
		expect(wb.sheets[0]!.scenarios?.list.map((x) => x.name)).toEqual(['Higher']);
		s.deleteScenario(0, 0);
		expect(wb.sheets[0]!.scenarios).toBeUndefined();
		s.undo();
		expect(wb.sheets[0]!.scenarios?.list).toHaveLength(1);
	});

	it('round-trips the fixture scenarios', async () => {
		const wb = await loadXlsx(new Uint8Array(readFileSync(FIXTURE)));
		const set = wb.sheets[0]!.scenarios;
		expect(set?.list.map((x) => x.name)).toEqual(['Low', 'High']);
		expect(set?.list[0]).toMatchObject({ user: 'Tester', comment: 'Pessimistic' });
		expect(set).toMatchObject({ current: 0, show: 0 });
		const again = await loadXlsx(await saveXlsx(wb));
		expect(again.sheets[0]!.scenarios).toEqual(set);
	});
});
