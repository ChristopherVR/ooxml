import { describe, expect, it } from 'vitest';
import { MAX_ROW, parseRange } from '../address.js';
import { getCell } from '../cells.js';
import { createWorkbook } from '../workbook.js';
import { createEditSession } from './session.js';

const range = (ref: string) => {
	const result = parseRange(ref);
	if (!result) throw new Error(ref);
	return result;
};

describe('paste into a selection', () => {
	it('repeats a single cell with one undo/redo step', () => {
		const workbook = createWorkbook();
		const session = createEditSession(workbook);
		const sheet = workbook.sheets[0]!;
		session.setCellValue(0, 0, 0, 'Monday');
		const payload = session.copy(0, range('A1'));
		expect(session.paste(0, range('C3:D4'), payload)).toEqual(range('C3:D4'));
		for (const row of [2, 3])
			for (const col of [2, 3]) expect(getCell(sheet, row, col)?.value).toBe('Monday');
		session.undo();
		for (const row of [2, 3])
			for (const col of [2, 3]) expect(getCell(sheet, row, col)).toBeUndefined();
		session.redo();
		expect(getCell(sheet, 3, 3)?.value).toBe('Monday');
	});

	it('tiles formulas and blanks with references relative to each tile', () => {
		const workbook = createWorkbook();
		const session = createEditSession(workbook, { recalc: false });
		const sheet = workbook.sheets[0]!;
		session.setCellInput(0, 0, 0, '=B1+$C$1');
		const payload = session.copy(0, range('A1:B1'));
		session.setRangeValues(0, { row: 2, col: 2 }, [
			['old', 'old', 'old', 'old'],
			['old', 'old', 'old', 'old'],
		]);
		session.paste(0, range('C3:F4'), payload);
		expect(getCell(sheet, 2, 2)?.formula).toBe('D3+$C$1');
		expect(getCell(sheet, 3, 4)?.formula).toBe('F4+$C$1');
		expect(getCell(sheet, 3, 5)).toBeUndefined();
		session.undo();
		expect(getCell(sheet, 3, 5)?.value).toBe('old');
	});

	it('rejects incompatible selections before changing any cells', () => {
		const workbook = createWorkbook();
		const session = createEditSession(workbook);
		session.setCellValue(0, 0, 0, 1);
		const payload = session.copy(0, range('A1:B2'));
		const label = session.undoLabel();
		expect(() => session.paste(0, range('C3:E5'), payload)).toThrow(RangeError);
		expect(session.undoLabel()).toBe(label);
		expect(getCell(workbook.sheets[0]!, 2, 2)).toBeUndefined();
	});

	it('tiles transposed text blocks', () => {
		const workbook = createWorkbook();
		const session = createEditSession(workbook);
		session.paste(0, range('C3:D6'), '1\t2', 'transpose');
		const sheet = workbook.sheets[0]!;
		expect([2, 3, 4, 5].map((row) => getCell(sheet, row, 3)?.value)).toEqual([1, 2, 1, 2]);
	});

	it('tiles merges and restores destination merges through undo and redo', () => {
		const workbook = createWorkbook();
		const session = createEditSession(workbook);
		const sheet = workbook.sheets[0]!;
		session.setCellValue(0, 0, 0, 'merged');
		session.merge(0, range('A1:B1'), 'merge');
		const payload = session.copy(0, range('A1:B1'));
		session.merge(0, range('C3:D4'), 'merge');
		session.paste(0, range('C3:F4'), payload);
		const expected = ['A1:B1', 'C3:D3', 'E3:F3', 'C4:D4', 'E4:F4'].map(range);
		expect(sheet.merges).toEqual(expected);
		session.undo();
		expect(sheet.merges).toEqual(['A1:B1', 'C3:D4'].map(range));
		session.redo();
		expect(sheet.merges).toEqual(expected);
	});

	it('preserves destination formatting in repeated values-only pastes', () => {
		const workbook = createWorkbook();
		const session = createEditSession(workbook);
		const sheet = workbook.sheets[0]!;
		session.setCellValue(0, 0, 0, 7);
		session.applyStyle(0, [range('C3:D4')], { numFmt: '0.00' });
		const styleId = getCell(sheet, 3, 3)?.styleId;
		session.paste(0, range('C3:D4'), session.copy(0, range('A1')), 'values');
		expect(getCell(sheet, 3, 3)).toMatchObject({ value: 7, styleId });
	});

	it('rolls back earlier tiles if a later tile extends beyond the worksheet', () => {
		const workbook = createWorkbook();
		const session = createEditSession(workbook);
		const sheet = workbook.sheets[0]!;
		session.setCellValue(0, MAX_ROW, 0, 'old');
		const before = session.undoLabel();
		expect(() =>
			session.paste(
				0,
				{
					start: { row: MAX_ROW, col: 0 },
					end: { row: MAX_ROW + 1, col: 0 },
				},
				'new',
			),
		).toThrow(RangeError);
		expect(getCell(sheet, MAX_ROW, 0)?.value).toBe('old');
		expect(session.undoLabel()).toBe(before);
	});
});
