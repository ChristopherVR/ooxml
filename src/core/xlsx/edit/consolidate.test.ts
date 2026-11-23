import { describe, expect, it } from 'vitest';
import { parseRange } from '../address';
import { getCell } from '../cells';
import { createWorkbook } from '../workbook';
import { createEditSession } from './session';
import type { SubtotalFunction, SubtotalFunctionCode } from './subtotal-fn';

describe('consolidateRanges', () => {
	it('consolidates same-sheet and cross-sheet ranges by position', () => {
		const workbook = createWorkbook();
		const session = createEditSession(workbook, { autoRowHeight: false });
		const s2 = session.addSheet('Q2');
		const dest = session.addSheet('Total');

		session.setRangeValues(0, { row: 0, col: 0 }, [
			[10, 20],
			[30, 40],
		]);
		session.setRangeValues(s2, { row: 0, col: 0 }, [
			[5, 15, 25],
			[35, 45, 55],
		]);

		const res = session.consolidate(
			dest,
			{ row: 1, col: 1 },
			[
				{ sheet: 0, range: parseRange('A1:B2')! },
				{ sheet: s2, range: parseRange('A1:C2')! },
			],
			{ fn: 'sum' },
		);

		expect(res).toEqual({ range: parseRange('B2:D3'), rows: 2, cols: 3 });
		const target = workbook.sheets[dest]!;
		expect(getCell(target, 1, 1)?.value).toBe(15);
		expect(getCell(target, 1, 2)?.value).toBe(35);
		expect(getCell(target, 1, 3)?.value).toBe(25);
		expect(getCell(target, 2, 1)?.value).toBe(65);
		expect(getCell(target, 2, 2)?.value).toBe(85);
		expect(getCell(target, 2, 3)?.value).toBe(55);
	});

	it.each<[SubtotalFunction, SubtotalFunctionCode, number]>([
		['average', 1, 20],
		['count', 2, 3],
		['counta', 3, 3],
		['max', 4, 30],
		['min', 5, 10],
		['product', 6, 6000],
		['stdev', 7, 10],
		['stdevp', 8, Math.sqrt(200 / 3)],
		['sum', 9, 60],
		['var', 10, 100],
		['varp', 11, 200 / 3],
	])('supports %s (code %s) by name and numeric code', (name, code, expected) => {
		for (const fn of [name, code] as const) {
			const workbook = createWorkbook();
			const session = createEditSession(workbook, { autoRowHeight: false });
			session.setRangeValues(0, { row: 0, col: 0 }, [[10], [20], [30]]);
			session.consolidate(
				0,
				{ row: 0, col: 2 },
				[
					{ sheet: 0, range: parseRange('A1:A1')! },
					{ sheet: 0, range: parseRange('A2:A2')! },
					{ sheet: 0, range: parseRange('A3:A3')! },
				],
				{ fn },
			);
			expect(getCell(workbook.sheets[0]!, 0, 2)?.value).toBeCloseTo(expected);
		}
	});

	it('distinguishes COUNT and COUNTA on mixed text/number cells and propagates errors', () => {
		const workbook = createWorkbook();
		const session = createEditSession(workbook, { autoRowHeight: false });
		session.setRangeValues(0, { row: 0, col: 0 }, [['abc'], [10], [true]]);

		const sources = [
			{ sheet: 0, range: parseRange('A1:A1')! },
			{ sheet: 0, range: parseRange('A2:A2')! },
			{ sheet: 0, range: parseRange('A3:A3')! },
		];
		session.consolidate(0, { row: 0, col: 2 }, sources, { fn: 'count' });
		session.consolidate(0, { row: 0, col: 3 }, sources, { fn: 'counta' });
		expect(getCell(workbook.sheets[0]!, 0, 2)?.value).toBe(1);
		expect(getCell(workbook.sheets[0]!, 0, 3)?.value).toBe(3);

		// STDEV on a single numeric value produces #DIV/0!
		session.consolidate(0, { row: 0, col: 4 }, [{ sheet: 0, range: parseRange('A2:A2')! }], {
			fn: 'stdev',
		});
		expect(getCell(workbook.sheets[0]!, 0, 4)?.value).toEqual({ error: '#DIV/0!' });
	});

	it('consolidates by category with both useTopRow and useLeftColumn (case-insensitive)', () => {
		const workbook = createWorkbook();
		const session = createEditSession(workbook, { autoRowHeight: false });
		const s2 = session.addSheet('Sheet2');
		const dest = session.addSheet('Consolidated');

		session.setRangeValues(0, { row: 0, col: 0 }, [
			['Product', 'Jan', 'Feb'],
			['Apples', 10, 20],
			['Oranges', 30, 40],
		]);
		session.setRangeValues(s2, { row: 0, col: 0 }, [
			['Product', 'feb', 'MAR'],
			['oranges', 5, 15],
			['Bananas', 25, 35],
			['APPLES', 100, 200],
		]);

		const res = session.consolidate(
			dest,
			{ row: 0, col: 0 },
			[
				{ sheet: 0, range: parseRange('A1:C3')! },
				{ sheet: s2, range: parseRange('A1:C4')! },
			],
			{ fn: 'sum', useTopRow: true, useLeftColumn: true },
		);

		expect(res).toEqual({ range: parseRange('A1:D4'), rows: 4, cols: 4 });
		const target = workbook.sheets[dest]!;

		// Header row preserves first-seen casing
		expect(getCell(target, 0, 0)).toBeUndefined();
		expect(getCell(target, 0, 1)?.value).toBe('Jan');
		expect(getCell(target, 0, 2)?.value).toBe('Feb');
		expect(getCell(target, 0, 3)?.value).toBe('MAR');

		// Apples: Jan=10, Feb=20+100=120, MAR=200
		expect(getCell(target, 1, 0)?.value).toBe('Apples');
		expect(getCell(target, 1, 1)?.value).toBe(10);
		expect(getCell(target, 1, 2)?.value).toBe(120);
		expect(getCell(target, 1, 3)?.value).toBe(200);

		// Oranges: Jan=30, Feb=40+5=45, MAR=15
		expect(getCell(target, 2, 0)?.value).toBe('Oranges');
		expect(getCell(target, 2, 1)?.value).toBe(30);
		expect(getCell(target, 2, 2)?.value).toBe(45);
		expect(getCell(target, 2, 3)?.value).toBe(15);

		// Bananas: Jan=empty, Feb=25, MAR=35
		expect(getCell(target, 3, 0)?.value).toBe('Bananas');
		expect(getCell(target, 3, 1)).toBeUndefined();
		expect(getCell(target, 3, 2)?.value).toBe(25);
		expect(getCell(target, 3, 3)?.value).toBe(35);
	});

	it('consolidates by top row only and by left column only', () => {
		const workbook = createWorkbook();
		const session = createEditSession(workbook, { autoRowHeight: false });
		const sheet = workbook.sheets[0]!;

		// Top row only
		session.setRangeValues(0, { row: 0, col: 0 }, [
			['Q1', 'Q2'],
			[10, 20],
			[30, 40],
		]);
		session.setRangeValues(0, { row: 0, col: 3 }, [
			['q2', 'Q3'],
			[5, 15],
		]);
		const topRes = session.consolidate(
			0,
			{ row: 5, col: 0 },
			[
				{ sheet: 0, range: parseRange('A1:B3')! },
				{ sheet: 0, range: parseRange('D1:E2')! },
			],
			{ fn: 'sum', useTopRow: true },
		);
		expect(topRes).toEqual({ range: parseRange('A6:C8'), rows: 3, cols: 3 });
		expect(getCell(sheet, 5, 0)?.value).toBe('Q1');
		expect(getCell(sheet, 5, 1)?.value).toBe('Q2');
		expect(getCell(sheet, 5, 2)?.value).toBe('Q3');
		expect(getCell(sheet, 6, 0)?.value).toBe(10);
		expect(getCell(sheet, 6, 1)?.value).toBe(25);
		expect(getCell(sheet, 6, 2)?.value).toBe(15);
		expect(getCell(sheet, 7, 0)?.value).toBe(30);
		expect(getCell(sheet, 7, 1)?.value).toBe(40);
		expect(getCell(sheet, 7, 2)).toBeUndefined();

		// Left column only
		session.setRangeValues(0, { row: 10, col: 0 }, [
			['North', 100],
			['South', 200],
		]);
		session.setRangeValues(0, { row: 10, col: 3 }, [
			['south', 50],
			['East', 75],
		]);
		const leftRes = session.consolidate(
			0,
			{ row: 15, col: 0 },
			[
				{ sheet: 0, range: parseRange('A11:B12')! },
				{ sheet: 0, range: parseRange('D11:E12')! },
			],
			{ fn: 'sum', useLeftColumn: true },
		);
		expect(leftRes).toEqual({ range: parseRange('A16:B18'), rows: 3, cols: 2 });
		expect(getCell(sheet, 15, 0)?.value).toBe('North');
		expect(getCell(sheet, 15, 1)?.value).toBe(100);
		expect(getCell(sheet, 16, 0)?.value).toBe('South');
		expect(getCell(sheet, 16, 1)?.value).toBe(250);
		expect(getCell(sheet, 17, 0)?.value).toBe('East');
		expect(getCell(sheet, 17, 1)?.value).toBe(75);
	});

	it('handles overlapping destination and source safely and recalculates dependent formulas with undo/redo', () => {
		const workbook = createWorkbook();
		const session = createEditSession(workbook, { autoRowHeight: false });
		const sheet = workbook.sheets[0]!;
		session.setRangeValues(0, { row: 0, col: 0 }, [
			[10, 20],
			[30, 40],
		]);
		session.setCellInput(0, 5, 0, '=SUM(A1:B2)');
		expect(getCell(sheet, 5, 0)?.value).toBe(100);

		// Consolidate A1:B2 with itself directly back into A1
		session.consolidate(
			0,
			{ row: 0, col: 0 },
			[
				{ sheet: 0, range: parseRange('A1:B2')! },
				{ sheet: 0, range: parseRange('A1:B2')! },
			],
			{ fn: 'sum' },
		);
		expect(session.undoLabel()).toBe('Consolidate');
		expect(getCell(sheet, 0, 0)?.value).toBe(20);
		expect(getCell(sheet, 1, 1)?.value).toBe(80);
		expect(getCell(sheet, 5, 0)?.value).toBe(200);

		expect(session.undo()).toBe(true);
		expect(getCell(sheet, 0, 0)?.value).toBe(10);
		expect(getCell(sheet, 1, 1)?.value).toBe(40);
		expect(getCell(sheet, 5, 0)?.value).toBe(100);

		expect(session.redo()).toBe(true);
		expect(getCell(sheet, 0, 0)?.value).toBe(20);
		expect(getCell(sheet, 5, 0)?.value).toBe(200);
	});

	it('validates arguments', () => {
		const workbook = createWorkbook();
		const session = createEditSession(workbook, { autoRowHeight: false });
		expect(() => session.consolidate(0, { row: 0, col: 0 }, [], { fn: 'sum' })).toThrow(RangeError);
		expect(() =>
			session.consolidate(0, { row: -1, col: 0 }, [{ sheet: 0, range: parseRange('A1:A1')! }], {
				fn: 'sum',
			}),
		).toThrow(RangeError);
		expect(() =>
			session.consolidate(0, { row: 0, col: 0 }, [{ sheet: 99, range: parseRange('A1:A1')! }], {
				fn: 'sum',
			}),
		).toThrow(RangeError);
	});
});
