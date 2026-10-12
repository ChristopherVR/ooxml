import { expect, it } from 'vitest';
import { createWorkbook } from '../workbook';
import { createEditSession } from '../edit/session';
import { saveXlsx } from '../write/index';
import { loadXlsx } from '../read/index';
import { paginateWorksheet } from './print-layout';
import { printTitles } from './print-titles';
import { parseHeaderFooter, renderHeaderFooter, pageHeaderFooter } from './header-footer';
import { createPrintMergeIndex } from './print-merges';

it('stores native print titles, repeats them on each page, and follows source edits with undo', async () => {
	const wb = createWorkbook({ sheets: ['Report'] });
	const edit = createEditSession(wb);
	edit.setCellInput(0, 0, 0, 'Title');
	edit.setCellInput(0, 79, 3, 'Last');
	edit.setPrintTitles(0, { rows: { start: 0, end: 0 }, columns: { start: 0, end: 0 } });
	edit.setPageBreak(0, 'row', 40);
	edit.setPageBreak(0, 'col', 2);
	const plan = paginateWorksheet(wb, 0);
	expect(plan.pages).toHaveLength(4);
	for (const page of plan.pages) {
		expect(page.rows[0]).toBe(0);
		expect(page.columns[0]).toBe(0);
		expect(page.rows.filter((row) => row === 0)).toHaveLength(1);
	}
	const loaded = await loadXlsx(await saveXlsx(wb));
	expect(printTitles(loaded, 0)).toEqual({
		rows: { start: 0, end: 0 },
		columns: { start: 0, end: 0 },
	});
	edit.renameSheet(0, 'Renamed');
	expect(wb.definedNames[0]!.formula).toContain('Renamed!');
	edit.insertRows(0, 0, 2);
	expect(printTitles(wb, 0).rows).toEqual({ start: 2, end: 2 });
	edit.undo();
	expect(printTitles(wb, 0).rows).toEqual({ start: 0, end: 0 });
	edit.setPrintTitles(0, { rows: undefined });
	expect(printTitles(wb, 0).rows).toBeUndefined();
	edit.undo();
	expect(printTitles(wb, 0).rows).toEqual({ start: 0, end: 0 });
});

it('evaluates field codes, sections, escaped ampersands and localized date/time', () => {
	const result = renderHeaderFooter('&L&A && &F&C&D &T&RPage &P of &N', {
		page: 5,
		pages: 4,
		sheet: 'Report',
		file: 'Book.xlsx',
		date: new Date(2026, 9, 12, 13, 4),
		locale: 'en-AU',
	});
	expect(result.sections.left).toBe('Report & Book.xlsx');
	expect(result.sections.right).toBe('Page 5 of 4');
	// Excel 16.0.20430.20146 printed "Offset 72" for "Offset &P+2" on page 7.
	// Native oddHeader XML ended with "Offset &amp;P+2".
	// Keep the raw offset token unsupported rather than assume arithmetic behavior.
	expect(parseHeaderFooter('&P+2').unsupported).toEqual(['&P+2']);
	expect(result.sections.center).toContain('12/10/26');
	expect(result.sections.center).toMatch(/1:04/);
	expect(result.unsupported).toEqual([]);
	expect(parseHeaderFooter('&C&"Calibri,Bold"&12&BHello &G').unsupported).toEqual([
		'&"Calibri,Bold"',
		'&12',
		'&B',
		'&G',
	]);
	expect(
		pageHeaderFooter(
			{
				header: 'odd',
				evenHeader: 'even',
				firstHeader: 'first',
				differentFirst: true,
				differentOddEven: true,
			},
			1,
			5,
		).header,
	).toBe('first');
	expect(
		pageHeaderFooter({ header: 'odd', evenHeader: 'even', differentOddEven: true }, 2, 6).header,
	).toBe('even');
});

it('indexes a large merged grid once and performs constant-time cell lookups', () => {
	const rows = Array.from({ length: 250 }, (_, i) => i);
	const columns = Array.from({ length: 80 }, (_, i) => i);
	const ranges = rows.flatMap((row) =>
		Array.from({ length: 20 }, (_, i) => ({
			start: { row, col: i * 4 },
			end: { row, col: i * 4 + 1 },
		})),
	);
	const index = createPrintMergeIndex(ranges, rows, columns);
	let covered = 0;
	for (const row of rows) for (const col of columns) if (index.at(row, col)) covered++;
	expect(covered).toBe(10000);
	expect(index.stats).toEqual({ mergesVisited: 5000, cellsIndexed: 10000, lookups: 20000 });
});

it('preserves unsupported title definitions unless explicitly replaced', () => {
	const wb = createWorkbook();
	const edit = createEditSession(wb);
	edit.setDefinedName({
		name: '_xlnm.Print_Titles',
		localSheet: 0,
		formula: 'OFFSET(Sheet1!A1,0,0,2,3)',
	});
	expect(printTitles(wb, 0)).toEqual({ unsupported: true });
	expect(() => edit.setPrintTitles(0, { rows: { start: 0, end: 1 } })).toThrow('both');
	expect(wb.definedNames[0]!.formula).toContain('OFFSET');
});

it('accounts for repeat-title capacity and oversized columns in fit-to-page scaling', () => {
	const wb = createWorkbook();
	const edit = createEditSession(wb);
	const sheet = wb.sheets[0]!;
	edit.setCellInput(0, 0, 0, 'Title');
	edit.setCellInput(0, 8, 1, 'Body');
	sheet.columns = [{ min: 0, max: 1, width: 120 }];
	edit.setPrintTitles(0, { columns: { start: 0, end: 0 }, rows: { start: 0, end: 0 } });
	edit.setPageSetup(0, { fitToWidth: 1, fitToHeight: 1 });
	const layout = paginateWorksheet(wb, 0);
	expect(layout.pages).toHaveLength(1);
	expect(layout.scale).toBeLessThan(0.5);
	expect(layout.pages[0]!.width * layout.scale).toBeLessThanOrEqual(layout.contentWidth);
});
