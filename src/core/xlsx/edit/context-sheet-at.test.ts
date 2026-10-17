import { describe, expect, it } from 'vitest';
import { createWorkbook } from '../workbook';
import { sheetAt } from './context';

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
