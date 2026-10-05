import { describe, expect, it } from 'vitest';
import { loadXlsx } from '../read/load.js';
import { createWorkbook } from '../workbook.js';
import { saveXlsx } from '../write/save.js';
import { legacyPasswordHash, verifySheetPassword, verifyWorkbookPassword } from './protection.js';
import { createEditSession } from './session.js';

const setup = () => {
	const wb = createWorkbook({ sheets: ['Sheet1'] });
	return { wb, s: createEditSession(wb, { recalc: false }) };
};

describe('page setup and print options', () => {
	it('patches the page setup as one undo step', () => {
		const { wb, s } = setup();
		s.setPageSetup(0, { orientation: 'landscape', scale: 80 });
		expect(wb.sheets[0]?.pageSetup).toEqual({ orientation: 'landscape', scale: 80 });
		expect(s.undoLabel()).toBe('Page setup');
		s.setPageSetup(0, { scale: undefined });
		expect(wb.sheets[0]?.pageSetup).toEqual({ orientation: 'landscape' });
		s.undo();
		s.undo();
		expect(wb.sheets[0]?.pageSetup).toBeUndefined();
	});
	it('models print options and round-trips them', async () => {
		const { wb, s } = setup();
		s.setPrintOptions(0, { gridLines: true, headings: true });
		expect(s.undoLabel()).toBe('Print options');
		s.setPrintOptions(0, { headings: false, horizontalCentered: true });
		expect(wb.sheets[0]?.printOptions).toEqual({ gridLines: true, horizontalCentered: true });
		const back = await loadXlsx(await saveXlsx(wb));
		expect(back.sheets[0]?.printOptions).toEqual({ gridLines: true, horizontalCentered: true });
		expect(back.sheets[0]?.preserved.has('printOptions')).toBe(false);
		s.setPrintOptions(0, { gridLines: false, horizontalCentered: false });
		expect(wb.sheets[0]?.printOptions).toBeUndefined();
	});
});

describe('legacy password hash', () => {
	it('matches the values Excel (and openpyxl) compute', () => {
		expect(legacyPasswordHash('secret')).toBe('DAA7');
		expect(legacyPasswordHash('test')).toBe('CBEB');
		expect(legacyPasswordHash('a')).toBe('CE88');
		expect(legacyPasswordHash('password1')).toBe('E1AE');
		expect(legacyPasswordHash('Abc123!')).toBe('D54E');
		expect(legacyPasswordHash('')).toBe('CE4B');
	});
});

describe('sheet protection', () => {
	it('protects with a password, verifies and round-trips', async () => {
		const { wb, s } = setup();
		s.setSheetProtection(0, { sheet: true, allow: ['formatCells'] }, 'secret');
		expect(s.undoLabel()).toBe('Protect sheet');
		const sheet = wb.sheets[0];
		expect(sheet?.protection).toEqual({
			sheet: true,
			allow: ['formatCells'],
			passwordHash: 'DAA7',
		});
		expect(verifySheetPassword(sheet, 'secret')).toBe(true);
		expect(verifySheetPassword(sheet, 'nope')).toBe(false);
		const back = await loadXlsx(await saveXlsx(wb));
		expect(back.sheets[0]?.protection?.passwordHash).toBe('DAA7');
		expect(() => s.setSheetProtection(0, undefined, 'nope')).toThrow(/not correct/);
		s.setSheetProtection(0, undefined, 'secret');
		expect(sheet?.protection).toBeUndefined();
		expect(s.undoLabel()).toBe('Unprotect sheet');
		s.undo();
		expect(sheet?.protection?.sheet).toBe(true);
	});
	it('treats an unprotected sheet or one without a password as open', () => {
		const { wb, s } = setup();
		expect(verifySheetPassword(wb.sheets[0], 'x')).toBe(true);
		s.setSheetProtection(0, { sheet: true });
		expect(verifySheetPassword(wb.sheets[0], 'anything')).toBe(true);
	});
});

describe('workbook protection', () => {
	it('locks the structure with a password and undoes', async () => {
		const { wb, s } = setup();
		s.setWorkbookProtection(true, 'test');
		expect(wb.structureLocked).toBe(true);
		expect(wb.workbookPasswordHash).toBe('CBEB');
		expect(s.undoLabel()).toBe('Protect workbook');
		expect(() => s.addSheet()).toThrow();
		const back = await loadXlsx(await saveXlsx(wb));
		expect(back.structureLocked).toBe(true);
		expect(back.workbookPasswordHash).toBe('CBEB');
		expect(verifyWorkbookPassword(back, 'test')).toBe(true);
		expect(() => s.setWorkbookProtection(false, 'wrong')).toThrow(/not correct/);
		s.setWorkbookProtection(false, 'test');
		expect(wb.structureLocked).toBeUndefined();
		expect(wb.workbookPasswordHash).toBeUndefined();
		s.undo();
		expect(wb.workbookPasswordHash).toBe('CBEB');
		s.undo();
		expect(wb.structureLocked).toBeUndefined();
	});
});
