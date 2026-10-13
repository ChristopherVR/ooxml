import { expect, it } from 'vitest';
import { createWorkbook } from '../workbook';
import { createEditSession } from '../edit/session';
import { parseRange } from '../address';
import { loadXlsx } from '../read/index';
import { saveXlsx } from '../write/index';
import { paginateWorksheet } from './print-layout';
import { sheetPageBreaks } from './print-breaks';

const setup = () => {
	const wb = createWorkbook();
	const edit = createEditSession(wb);
	for (let r = 0; r < 10; r++) for (let c = 0; c < 4; c++) edit.setCellInput(0, r, c, `${r}:${c}`);
	return { wb, edit };
};

it('paginates manual row/column breaks in native down-then-over and over-then-down order', async () => {
	const { wb, edit } = setup();
	edit.setPageBreak(0, 'row', 5);
	edit.setPageBreak(0, 'col', 2);
	const layout = paginateWorksheet(wb, 0);
	expect(layout.pages.map((page) => page.range)).toEqual(
		['A1:B5', 'A6:B10', 'C1:D5', 'C6:D10'].map((ref) => parseRange(ref)),
	);
	edit.setPageSetup(0, {
		pageOrder: 'overThenDown',
		paperSize: 9,
		orientation: 'landscape',
		scale: 75,
	});
	expect(paginateWorksheet(wb, 0).pages.map((page) => page.range)).toEqual(
		['A1:B5', 'C1:D5', 'A6:B10', 'C6:D10'].map((ref) => parseRange(ref)),
	);
	const loaded = await loadXlsx(await saveXlsx(wb));
	expect(loaded.sheets[0]!.pageSetup).toMatchObject({
		pageOrder: 'overThenDown',
		paperSize: 9,
		orientation: 'landscape',
		scale: 75,
	});
	expect(sheetPageBreaks(loaded.sheets[0]!, 'row')[0]).toMatchObject({ before: 5, manual: true });
	expect(paginateWorksheet(loaded, 0).pages).toHaveLength(4);
});

it('fits complete rows and columns, ignores manual breaks, and handles hidden dimensions', () => {
	const { wb, edit } = setup();
	const sheet = wb.sheets[0]!;
	sheet.columns = [{ min: 0, max: 3, width: 65 }];
	edit.setPageBreak(0, 'col', 1);
	edit.setPageSetup(0, { fitToWidth: 2, fitToHeight: 0 });
	const layout = paginateWorksheet(wb, 0);
	expect(layout.pages).toHaveLength(2);
	expect(layout.scale).toBeLessThan(1);
	expect(layout.warnings).toContain('Fit to pages ignores manual page breaks.');
	sheet.columns[0]!.hidden = true;
	expect(paginateWorksheet(wb, 0).pages).toEqual([]);
});

it('restores authored and reset breaks with undo and keeps unknown XML and automatic breaks', () => {
	const { wb, edit } = setup();
	const sheet = wb.sheets[0]!;
	sheet.preserved.set('rowBreaks', [
		'<rowBreaks xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" vendor="keep"><brk id="3" max="16383"/><brk id="5" max="16383" man="1" vendor="keep"/></rowBreaks>',
	]);
	edit.setPageBreak(0, 'row', 5);
	expect(sheet.preserved.get('rowBreaks')?.[0]).toContain('vendor="keep"');
	edit.resetPageBreaks(0);
	expect(sheetPageBreaks(sheet, 'row')).toEqual([{ before: 3, min: 0, max: 16383, manual: false }]);
	edit.undo();
	expect(sheetPageBreaks(sheet, 'row')).toHaveLength(2);
	edit.setPageBreak(0, 'row', 5, false);
	expect(sheetPageBreaks(sheet, 'row')).toHaveLength(1);
	edit.undo();
	edit.insertRows(0, 0, 2);
	expect(sheetPageBreaks(sheet, 'row').map((entry) => entry.before)).toEqual([5, 7]);
	expect(() => edit.setPageBreak(0, 'row', 0)).toThrow('after the first');
});

it('rejects impossible margins and bounds preview work', () => {
	const { wb, edit } = setup();
	edit.setPageSetup(0, {
		margins: { left: 10, right: 10, top: 0, bottom: 0, header: 0, footer: 0 },
	});
	expect(() => paginateWorksheet(wb, 0)).toThrow('Margins');
	edit.undo();
	edit.setPageSetup(0, { printArea: parseRange('A1:XFD1048576')! });
	expect(() => paginateWorksheet(wb, 0)).toThrow('one million');
});
