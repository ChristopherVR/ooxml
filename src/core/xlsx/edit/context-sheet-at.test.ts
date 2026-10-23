import { describe, expect, it } from 'vitest';
import { createWorkbook } from '../workbook';
import { sheetAt, sheetDrawingAt } from './context';

describe('sheetAt', () => {
	const workbook = createWorkbook();

	it('returns the sheet at a valid index', () => {
		expect(sheetAt(workbook, 0)).toBe(workbook.sheets[0]);
	});

	it('refuses indexes that name no sheet, including negative ones and prototype keys', () => {
		for (const index of [-1, workbook.sheets.length, 1.5, Number.NaN, '__proto__' as never]) {
			expect(() => sheetAt(workbook, index), String(index)).toThrow(RangeError);
		}
	});
});

describe('sheetDrawingAt', () => {
	it('refuses indexes that name no drawing, including negative ones and prototype keys', () => {
		const workbook = createWorkbook();
		for (const index of [-1, 0, 1.5, Number.NaN, '__proto__' as never, 'constructor' as never]) {
			expect(sheetDrawingAt(workbook, 0, index), String(index)).toBeUndefined();
		}
	});
});
