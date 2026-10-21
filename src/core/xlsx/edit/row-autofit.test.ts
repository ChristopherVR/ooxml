import { describe, expect, it } from 'vitest';
import { putCell } from '../cells';
import { approximateMeasure, autoFitRowHeight } from '../layout/row-autofit';
import { internStyle } from '../styles';
import { createWorkbook, defaultCellStyle } from '../workbook';
import { autoGrowRows } from './row-autofit';
import { testContext } from './test-context';

const ITERATORS = [
	Symbol.iterator,
	'entries',
	'keys',
	'values',
	'forEach',
	'some',
	'every',
	'find',
	'findIndex',
	'filter',
	'map',
	'reduce',
];

/**
 * Counts full passes over `target` (calls of its iteration methods). The wrappers are own,
 * non-enumerable properties, so the collection still clones like the original.
 */
function countPasses(target: object): { passes: number } {
	const counter = { passes: 0 };
	const proto = Object.getPrototypeOf(target) as Record<PropertyKey, unknown>;
	for (const name of ITERATORS) {
		const original = proto[name];
		if (typeof original !== 'function') continue;
		Object.defineProperty(target, name, {
			configurable: true,
			value(this: unknown, ...args: unknown[]) {
				counter.passes++;
				return original.apply(this, args);
			},
		});
	}
	return counter;
}

function setup() {
	const wb = createWorkbook();
	const sheet = wb.sheets[0];
	if (!sheet) throw new Error('no sheet');
	const base = defaultCellStyle();
	const big = internStyle(wb, { ...base, font: { ...base.font, size: 18 } });
	return { wb, sheet, big };
}

describe('autoGrowRows', () => {
	it('grows a row without a custom height, undoably', () => {
		const { wb, sheet, big } = setup();
		putCell(sheet, 2, 0, { value: 'Title', styleId: big });
		const ctx = testContext(wb);
		expect(autoGrowRows(ctx, 0, [2])).toEqual([2]);
		expect(sheet.rowInfo.get(2)?.height).toBe(23.25);
		expect(sheet.rowInfo.get(2)?.customHeight).toBeUndefined();
		ctx.undo();
		expect(sheet.rowInfo.get(2)).toBeUndefined();
	});

	it('shrinks back to the default and records nothing when unchanged', () => {
		const { wb, sheet } = setup();
		sheet.rowInfo.set(1, { height: 30 });
		putCell(sheet, 1, 0, { value: 'plain' });
		const ctx = testContext(wb);
		expect(autoGrowRows(ctx, 0, [1])).toEqual([1]);
		expect(sheet.rowInfo.has(1)).toBe(false);
		expect(autoGrowRows(ctx, 0, [1, 5])).toEqual([]);
		expect(ctx.steps).toHaveLength(1);
	});

	it('stores the height the one-row fit computes for every row, capped at 409 points', () => {
		const { wb, sheet, big } = setup();
		const wrap = internStyle(wb, { ...defaultCellStyle(), alignment: { wrapText: true } });
		for (let row = 0; row < 40; row++)
			putCell(sheet, row, row % 3, {
				value: `word${' word'.repeat(row * 4)}`,
				styleId: row % 2 ? wrap : big,
			});
		sheet.merges.push({ start: { row: 3, col: 0 }, end: { row: 4, col: 1 } });
		sheet.rowInfo.set(7, { height: 50, customHeight: true });
		const rows = Array.from({ length: 42 }, (_, row) => row);
		const expected = rows.map((row) => {
			if (row === 7) return 50;
			const fit = Math.min(409, autoFitRowHeight(wb, 0, row, approximateMeasure));
			return fit === sheet.defaultRowHeight ? undefined : fit;
		});
		expect(expected).toContain(409);
		autoGrowRows(testContext(wb), 0, rows);
		expect(rows.map((row) => sheet.rowInfo.get(row)?.height)).toEqual(expected);
	});

	it('reads the sheet-wide collections a fixed number of times, however many rows it fits', () => {
		const passes = (rowCount: number) => {
			const { wb, sheet, big } = setup();
			// Wrapped text, so the fit reads the column widths too.
			const wrap = internStyle(wb, { ...defaultCellStyle(), alignment: { wrapText: true } });
			for (let row = 0; row < rowCount; row++) {
				putCell(sheet, row, 0, { value: `Title ${row}`, styleId: big });
				putCell(sheet, row, 1, { value: row });
				putCell(sheet, row, 4, { value: `wrapped words in row ${row}`, styleId: wrap });
				if (row % 3 === 0) sheet.rowInfo.set(row, { height: 30 });
			}
			sheet.columns.push({ min: 0, max: 0, width: 20, customWidth: true });
			sheet.merges.push({ start: { row: 1, col: 2 }, end: { row: 2, col: 3 } });
			const counts = {
				rows: countPasses(sheet.rows),
				rowInfo: countPasses(sheet.rowInfo),
				columns: countPasses(sheet.columns),
				merges: countPasses(sheet.merges),
			};
			const rows = Array.from({ length: rowCount }, (_, row) => row);
			expect(autoGrowRows(testContext(wb), 0, rows)).toHaveLength(rowCount);
			return Object.fromEntries(Object.entries(counts).map(([name, c]) => [name, c.passes]));
		};
		const few = passes(10);
		expect(few.columns).toBeGreaterThan(0);
		expect(passes(400)).toEqual(few);
	});

	it('leaves custom-height and hidden rows alone', () => {
		const { wb, sheet, big } = setup();
		putCell(sheet, 0, 0, { value: 'x', styleId: big });
		putCell(sheet, 1, 0, { value: 'x', styleId: big });
		sheet.rowInfo.set(0, { height: 12, customHeight: true });
		sheet.rowInfo.set(1, { hidden: true });
		const ctx = testContext(wb);
		expect(autoGrowRows(ctx, 0, [0, 1])).toEqual([]);
		expect(sheet.rowInfo.get(0)?.height).toBe(12);
	});
});
