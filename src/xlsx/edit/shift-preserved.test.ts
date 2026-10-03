import { describe, expect, it } from 'vitest';
import { parseRange } from '../address.js';
import type { Workbook } from '../model.js';
import { createWorkbook } from '../workbook.js';
import { createEditSession } from './session.js';

const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const X14 = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main';
const XM = 'http://schemas.microsoft.com/office/excel/2006/main';

// The shapes Excel 16 writes (review-edit x3.ps1): a manual break before row 6, a sparkline in E1
// over A1:A10, an x14 data bar on C1:C10 and an x14 list validation on G2:G4.
const ROW_BREAKS = `<rowBreaks count="1" manualBreakCount="1" xmlns="${MAIN}"><brk id="5" max="16383" man="1"/></rowBreaks>`;
const COL_BREAKS = `<colBreaks count="1" manualBreakCount="1" xmlns="${MAIN}"><brk id="3" max="1048575" man="1"/></colBreaks>`;
const EXT_LST =
	`<extLst xmlns="${MAIN}">` +
	`<ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}" xmlns:x14="${X14}"><x14:conditionalFormattings>` +
	`<x14:conditionalFormatting xmlns:xm="${XM}"><x14:cfRule type="dataBar" id="{A}"><x14:dataBar/></x14:cfRule><xm:sqref>C1:C10</xm:sqref></x14:conditionalFormatting>` +
	`</x14:conditionalFormattings></ext>` +
	`<ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}" xmlns:x14="${X14}"><x14:dataValidations count="1" xmlns:xm="${XM}">` +
	`<x14:dataValidation type="list"><x14:formula1><xm:f>Sheet2!$A$1:$A$3</xm:f></x14:formula1><xm:sqref>G2:G4</xm:sqref></x14:dataValidation>` +
	`</x14:dataValidations></ext>` +
	`<ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}" xmlns:x14="${X14}"><x14:sparklineGroups xmlns:xm="${XM}">` +
	`<x14:sparklineGroup><x14:colorSeries rgb="FF376092"/><x14:sparklines><x14:sparkline><xm:f>Sheet1!A1:A10</xm:f><xm:sqref>E1</xm:sqref></x14:sparkline></x14:sparklines></x14:sparklineGroup>` +
	`</x14:sparklineGroups></ext></extLst>`;

function setup(): { wb: Workbook; s: ReturnType<typeof createEditSession> } {
	const wb = createWorkbook({ sheets: ['Sheet1', 'Sheet2'] });
	const sheet = wb.sheets[0]!;
	sheet.preserved.set('rowBreaks', [ROW_BREAKS]);
	sheet.preserved.set('colBreaks', [COL_BREAKS]);
	sheet.preserved.set('extLst', [EXT_LST]);
	return { wb, s: createEditSession(wb, { recalc: false }) };
}
const xml = (wb: Workbook, key: string, sheet = 0) => wb.sheets[sheet]!.preserved.get(key)?.[0];

describe('preserved worksheet XML follows row and column edits', () => {
	// Excel 16: inserting a row above row 1 moves the break to before row 7, the sparkline to E2
	// over A2:A11 and the data bar to C2:C11 (review-edit x4.ps1 after the fix).
	it('shifts page breaks, sparklines and x14 ranges on a row insert', () => {
		const { wb, s } = setup();
		s.insertRows(0, 0, 1);
		expect(xml(wb, 'rowBreaks')).toContain('<brk id="6"');
		expect(xml(wb, 'colBreaks')).toBe(COL_BREAKS);
		const ext = xml(wb, 'extLst')!;
		expect(ext).toContain('<xm:sqref>C2:C11</xm:sqref>');
		expect(ext).toContain('<xm:f>Sheet1!A2:A11</xm:f><xm:sqref>E2</xm:sqref>');
		expect(ext).toContain('<xm:sqref>G3:G5</xm:sqref>');
		expect(ext).toContain('<xm:f>Sheet2!$A$1:$A$3</xm:f>');
	});

	it('shifts column breaks and drops a sparkline whose cell is deleted', () => {
		const { wb, s } = setup();
		s.deleteColumns(0, 4, 1);
		expect(xml(wb, 'colBreaks')).toBe(COL_BREAKS);
		const ext = xml(wb, 'extLst')!;
		expect(ext).not.toContain('sparkline');
		expect(ext).toContain('<xm:sqref>F2:F4</xm:sqref>');
		s.insertColumns(0, 0, 2);
		expect(xml(wb, 'colBreaks')).toContain('<brk id="5"');
	});

	// Excel 16: deleting the row a manual break sits before removes the break.
	it('removes a break whose row is deleted and clips x14 ranges', () => {
		const { wb, s } = setup();
		s.deleteRows(0, 5, 1);
		expect(xml(wb, 'rowBreaks')).toBeUndefined();
		const ext = xml(wb, 'extLst')!;
		expect(ext).toContain('<xm:sqref>C1:C9</xm:sqref>');
		expect(ext).toContain('<xm:f>Sheet1!A1:A9</xm:f>');
	});

	it('drops a validation deleted entirely and updates its count', () => {
		const { wb, s } = setup();
		s.deleteRows(0, 1, 3);
		const ext = xml(wb, 'extLst')!;
		expect(ext).not.toContain('dataValidation ');
		expect(ext).not.toContain('CCE6A557');
		expect(xml(wb, 'rowBreaks')).toContain('<brk id="2"');
	});

	it('rewrites sparkline sources on another sheet and keeps breaks for cell shifts', () => {
		const { wb, s } = setup();
		wb.sheets[1]!.preserved.set('extLst', [EXT_LST.replace('Sheet1!A1:A10', 'Sheet1!B1:B10')]);
		s.insertCellsShift(0, parseRange('A1:B1')!, 'down');
		expect(xml(wb, 'rowBreaks')).toBe(ROW_BREAKS);
		expect(xml(wb, 'extLst', 1)).toContain('<xm:f>Sheet1!B2:B11</xm:f>');
		expect(xml(wb, 'extLst', 1)).toContain('<xm:sqref>E1</xm:sqref>');
		expect(xml(wb, 'extLst')).toContain('<xm:sqref>E1</xm:sqref>');
	});

	it('undoes the shift', () => {
		const { wb, s } = setup();
		s.insertRows(0, 0, 2);
		s.undo();
		expect(xml(wb, 'rowBreaks')).toBe(ROW_BREAKS);
		expect(xml(wb, 'extLst')).toBe(EXT_LST);
	});
});
