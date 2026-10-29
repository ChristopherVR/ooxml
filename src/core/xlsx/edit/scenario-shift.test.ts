import { describe, expect, it } from 'vitest';
import { getCell } from '../cells';
import { createWorkbook } from '../workbook';
import { createEditSession } from './session';

// Changing cells B2 and C2 with two scenarios; results in B5:C5 (sqref).
function setup() {
	const wb = createWorkbook({ sheets: ['Sheet1'] });
	const s = createEditSession(wb);
	s.setCellValue(0, 1, 1, 10);
	s.setCellValue(0, 1, 2, 20);
	s.addScenario(0, { name: 'Low', cells: 'B2,C2', values: { B2: '1', C2: '2' } });
	s.addScenario(0, { name: 'High', cells: 'B2,C2', values: { B2: '100', C2: '200' } });
	const sheet = wb.sheets[0]!;
	sheet.scenarios!.sqref = 'B5:C5 E1';
	return { wb, s, sheet };
}
const refs = (sheet: ReturnType<typeof setup>['sheet'], i = 0) =>
	sheet.scenarios!.list[i]!.cells.map((c) => c.ref);

describe('scenarios on row and column edits', () => {
	it('moves changing cells and result cells with a row insert', () => {
		const { wb, s, sheet } = setup();
		s.insertRows(0, 0, 2);
		expect(refs(sheet)).toEqual(['B4', 'C4']);
		expect(refs(sheet, 1)).toEqual(['B4', 'C4']);
		expect(sheet.scenarios!.sqref).toBe('B7:C7 E3');
		// Showing a scenario now writes into the moved cells.
		s.showScenario(0, 1);
		expect(getCell(wb.sheets[0]!, 3, 1)?.value).toBe(100);
		expect(getCell(wb.sheets[0]!, 1, 1)?.value ?? null).toBe(null);
		s.undo();
		s.undo();
		expect(refs(sheet)).toEqual(['B2', 'C2']);
		expect(sheet.scenarios!.sqref).toBe('B5:C5 E1');
	});

	it('moves changing cells with a column insert and leaves earlier ones alone', () => {
		const { s, sheet } = setup();
		s.insertColumns(0, 2, 1);
		expect(refs(sheet)).toEqual(['B2', 'D2']);
		expect(sheet.scenarios!.sqref).toBe('B5:D5 F1');
		s.insertRows(0, 10, 3);
		expect(refs(sheet)).toEqual(['B2', 'D2']);
	});

	it('marks a deleted changing cell and skips it when showing the scenario', () => {
		const { wb, s, sheet } = setup();
		s.deleteColumns(0, 2, 1);
		expect(sheet.scenarios!.list[0]!.cells.map((c) => [c.ref, c.deleted ?? false])).toEqual([
			['B2', false],
			['C2', true],
		]);
		s.showScenario(0, 1);
		expect(getCell(wb.sheets[0]!, 1, 1)?.value).toBe(100);
		expect(getCell(wb.sheets[0]!, 1, 2)?.value ?? null).toBe(null);
	});

	it('trims and drops result ranges on deletes', () => {
		const { s, sheet } = setup();
		s.deleteRows(0, 4, 1);
		expect(sheet.scenarios!.sqref).toBe('E1');
		s.deleteRows(0, 0, 1);
		expect(sheet.scenarios!.sqref).toBeUndefined();
	});
});
