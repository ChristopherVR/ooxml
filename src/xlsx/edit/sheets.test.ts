import { describe, expect, it } from 'vitest';
import { parseAddress, parseRange } from '../address.js';
import { getCell } from '../cells.js';
import type { Workbook } from '../model.js';
import { createWorkbook } from '../workbook.js';
import { createEditSession } from './session.js';

const A = (ref: string) => {
	const a = parseAddress(ref);
	if (!a) throw new Error(ref);
	return a;
};
const R = (ref: string) => {
	const r = parseRange(ref);
	if (!r) throw new Error(ref);
	return r;
};
const cell = (wb: Workbook, ref: string, i = 0) => {
	const ws = wb.sheets[i];
	return ws && getCell(ws, A(ref).row, A(ref).col);
};
const names = (wb: Workbook) => wb.sheets.map((s) => s.name);
const setup = (sheets = ['Sheet1', 'Sheet2', 'Sheet3']) => {
	const wb = createWorkbook({ sheets });
	return { wb, s: createEditSession(wb, { recalc: false }) };
};

describe('add and delete sheets', () => {
	it('appends a sheet with the next free name and makes it active', () => {
		const { wb, s } = setup();
		expect(s.addSheet()).toBe(3);
		expect(names(wb)).toEqual(['Sheet1', 'Sheet2', 'Sheet3', 'Sheet4']);
		expect(wb.activeSheet).toBe(3);
		expect(wb.sheets[3]?.sheetId).toBe(4);
		s.undo();
		expect(names(wb)).toHaveLength(3);
		expect(wb.activeSheet).toBe(0);
	});
	it('inserts at a position and shifts sheet-scoped names', () => {
		const { wb, s } = setup();
		s.setDefinedName({ name: 'Local', formula: 'Sheet2!$A$1', localSheet: 1 });
		s.addSheet('Data', 0);
		expect(names(wb)[0]).toBe('Data');
		expect(wb.definedNames[0]?.localSheet).toBe(2);
	});
	it('rejects invalid or duplicate names', () => {
		const { s } = setup();
		expect(() => s.addSheet('sheet1')).toThrow(/taken/);
		expect(() => s.addSheet('a/b')).toThrow();
		expect(() => s.addSheet('x'.repeat(32))).toThrow();
	});
	it('deletes a sheet with its local names and fixes the active sheet', () => {
		const { wb, s } = setup();
		s.setDefinedName({ name: 'Local', formula: 'Sheet2!$A$1', localSheet: 1 });
		s.setDefinedName({ name: 'Later', formula: 'Sheet3!$A$1', localSheet: 2 });
		wb.activeSheet = 2;
		s.deleteSheet(1);
		expect(names(wb)).toEqual(['Sheet1', 'Sheet3']);
		expect(wb.definedNames).toEqual([{ name: 'Later', formula: 'Sheet3!$A$1', localSheet: 1 }]);
		expect(wb.activeSheet).toBe(1);
		s.undo();
		expect(names(wb)).toEqual(['Sheet1', 'Sheet2', 'Sheet3']);
		expect(wb.definedNames).toHaveLength(2);
	});
	it('refuses to delete the last visible sheet', () => {
		const { s } = setup(['Only']);
		expect(() => s.deleteSheet(0)).toThrow(/visible/);
	});
	it('honours workbook structure protection', () => {
		const { wb, s } = setup();
		wb.structureLocked = true;
		expect(() => s.addSheet()).toThrow(/protected/);
		expect(() => s.renameSheet(0, 'X')).toThrow(/protected/);
	});
});

describe('rename', () => {
	it('renames and updates formulas on every sheet', () => {
		const { wb, s } = setup();
		s.setCellInput(1, 0, 0, '=Sheet1!A1+Sheet1!B2');
		s.setCellInput(0, 0, 1, '=A1');
		s.renameSheet(0, 'Data');
		expect(names(wb)[0]).toBe('Data');
		expect(cell(wb, 'A1', 1)?.formula).toBe('Data!A1+Data!B2');
		expect(cell(wb, 'B1', 0)?.formula).toBe('A1');
		s.undo();
		expect(cell(wb, 'A1', 1)?.formula).toBe('Sheet1!A1+Sheet1!B2');
	});
	it('quotes names that need it', () => {
		const { wb, s } = setup();
		s.setCellInput(1, 0, 0, '=Sheet1!A1');
		s.renameSheet(0, 'My Data');
		expect(cell(wb, 'A1', 1)?.formula).toBe("'My Data'!A1");
	});
	it('updates defined names, validations and hyperlink locations', () => {
		const { wb, s } = setup();
		s.setDefinedName({ name: 'Total', formula: 'Sheet1!$B$1' });
		s.setDataValidation(1, { ranges: [], type: 'list', formula1: 'Sheet1!$A$1:$A$3' }, R('A1'));
		s.setHyperlink(1, R('B1'), { location: 'Sheet1!A1' });
		s.renameSheet(0, 'Inputs');
		expect(wb.definedNames[0]?.formula).toBe('Inputs!$B$1');
		expect(wb.sheets[1]?.dataValidations[0]?.formula1).toBe('Inputs!$A$1:$A$3');
		expect(wb.sheets[1]?.hyperlinks[0]?.location).toBe('Inputs!A1');
	});
	it('allows changing only the case and ignores a no-op rename', () => {
		const { wb, s } = setup();
		s.renameSheet(0, 'SHEET1');
		expect(names(wb)[0]).toBe('SHEET1');
		s.renameSheet(0, 'SHEET1');
		expect(s.undoLabel()).toBe('Rename sheet');
		s.undo();
		expect(s.canUndo()).toBe(false);
	});
});

describe('move, duplicate, hide, colour', () => {
	it('moves a sheet and keeps the active sheet and local names attached', () => {
		const { wb, s } = setup();
		s.setDefinedName({ name: 'L', formula: 'Sheet1!$A$1', localSheet: 0 });
		wb.activeSheet = 0;
		s.moveSheet(0, 2);
		expect(names(wb)).toEqual(['Sheet2', 'Sheet3', 'Sheet1']);
		expect(wb.activeSheet).toBe(2);
		expect(wb.definedNames[0]?.localSheet).toBe(2);
		s.moveSheet(2, 1);
		expect(names(wb)).toEqual(['Sheet2', 'Sheet1', 'Sheet3']);
		s.undo();
		s.undo();
		expect(names(wb)).toEqual(['Sheet1', 'Sheet2', 'Sheet3']);
	});
	it('duplicates a sheet with its content after the original', () => {
		const { wb, s } = setup();
		s.setCellValue(0, 0, 0, 'x');
		s.setRangeValues(0, A('B1'), [['h'], [1]]);
		s.createTable(0, R('B1:B2'), true);
		const index = s.duplicateSheet(0);
		expect(index).toBe(1);
		expect(names(wb)).toEqual(['Sheet1', 'Sheet1 (2)', 'Sheet2', 'Sheet3']);
		expect(cell(wb, 'A1', 1)?.value).toBe('x');
		expect(wb.sheets[1]?.sheetId).toBe(4);
		expect(wb.sheets[1]?.tables[0]?.name).toBe('Table1_2');
		expect(wb.sheets[1]?.tables[0]?.id).toBe(2);
		s.setCellValue(1, 0, 0, 'changed');
		expect(cell(wb, 'A1', 0)?.value).toBe('x');
	});
	it('hides and unhides sheets but keeps one visible', () => {
		const { wb, s } = setup(['A', 'B']);
		wb.activeSheet = 0;
		s.setSheetState(0, 'hidden');
		expect(wb.sheets[0]?.state).toBe('hidden');
		expect(wb.activeSheet).toBe(1);
		expect(() => s.setSheetState(1, 'veryHidden')).toThrow(/visible/);
		s.setSheetState(0, 'visible');
		expect(wb.sheets[0]?.state).toBe('visible');
		s.undo();
		s.undo();
		expect(wb.activeSheet).toBe(0);
	});
	it('sets and clears the tab colour', () => {
		const { wb, s } = setup();
		s.setTabColor(0, { rgb: 'FF0000' });
		expect(wb.sheets[0]?.tabColor).toEqual({ rgb: 'FF0000' });
		s.setTabColor(0);
		expect(wb.sheets[0]?.tabColor).toBeUndefined();
		s.undo();
		expect(wb.sheets[0]?.tabColor).toEqual({ rgb: 'FF0000' });
	});
});
