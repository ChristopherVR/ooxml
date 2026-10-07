import { describe, expect, it } from 'vitest';
import { parseRange } from '../address';
import { getCell } from '../cells';
import { styleAt } from '../styles';
import { createWorkbook } from '../workbook';
import { createEditSession } from './session';

const range = (ref: string) => {
	const result = parseRange(ref);
	if (!result) throw new Error(ref);
	return result;
};

describe('copy fill mode', () => {
	it.each(['down', 'right'] as const)(
		'repeats dates and text %s without extending a series',
		(axis) => {
			const workbook = createWorkbook();
			const session = createEditSession(workbook);
			const sheet = workbook.sheets[0]!;
			session.setCellInput(0, 0, 0, '2026-10-07');
			session.applyStyle(0, [range('A1')], { numFmt: 'yyyy-mm-dd', font: { bold: true } });
			const original = structuredClone(getCell(sheet, 0, 0));
			session.fill(0, range('A1'), range(axis === 'down' ? 'A1:A3' : 'A1:C1'), 'copy');
			for (let i = 1; i <= 2; i++)
				expect(getCell(sheet, axis === 'down' ? i : 0, axis === 'right' ? i : 0)).toEqual(original);
			session.undo();
			expect(getCell(sheet, axis === 'down' ? 1 : 0, axis === 'right' ? 1 : 0)).toBeUndefined();
			session.redo();
			expect(
				styleAt(
					workbook,
					getCell(sheet, axis === 'down' ? 1 : 0, axis === 'right' ? 1 : 0)?.styleId,
				).font.bold,
			).toBe(true);
			session.setCellValue(0, 0, 0, 'Monday');
			session.fill(0, range('A1'), range(axis === 'down' ? 'A1:A3' : 'A1:C1'), 'copy');
			expect(getCell(sheet, axis === 'down' ? 2 : 0, axis === 'right' ? 2 : 0)?.value).toBe(
				'Monday',
			);
		},
	);

	it('still translates relative and mixed references in copy mode', () => {
		const workbook = createWorkbook();
		const session = createEditSession(workbook, { recalc: false });
		session.setCellInput(0, 0, 0, '=B1+$C1+D$1+$E$1');
		session.fill(0, range('A1'), range('A1:A3'), 'copy');
		expect(getCell(workbook.sheets[0]!, 2, 0)?.formula).toBe('B3+$C3+D$1+$E$1');
	});

	it('keeps default AutoFill series behavior', () => {
		const workbook = createWorkbook();
		const session = createEditSession(workbook);
		session.setCellValue(0, 0, 0, 'Monday');
		session.fill(0, range('A1'), range('A1:A3'));
		expect(getCell(workbook.sheets[0]!, 2, 0)?.value).toBe('Wednesday');
	});
});
