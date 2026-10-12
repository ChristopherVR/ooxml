import { expect, it, vi } from 'vitest';
import { createWorkbook, createEditSession } from 'ooxml-core/xlsx';
import { buildPrintHtml } from './print-pages';

it('renders separate physical pages from the core plan with settings, headings and escaping', () => {
	const wb = createWorkbook();
	const edit = createEditSession(wb);
	edit.setCellInput(0, 0, 0, '<script>unsafe</script>');
	edit.setCellInput(0, 8, 3, 'Last');
	edit.setPageBreak(0, 'row', 5);
	edit.setPageBreak(0, 'col', 2);
	edit.setPageSetup(0, { orientation: 'landscape', paperSize: 9, scale: 75 });
	edit.setPrintOptions(0, { gridLines: true, headings: true });
	const html = buildPrintHtml(wb, 0, 'Book<1>');
	expect(html.match(/data-page="/g)).toHaveLength(4);
	expect(html).toContain('transform:scale(0.75)');
	expect(html).toContain('Book&#60;1&#62;');
	expect(html).toContain('&#60;script&#62;unsafe&#60;/script&#62;');
	expect(html).not.toContain('<script>');
	expect(html).toContain('<th>A</th>');
	expect(html).toContain('1px solid #d9d9d9');
});

it('renders indexed merges without per-cell merge scans on a large text sheet', () => {
	const wb = createWorkbook();
	const edit = createEditSession(wb);
	const sheet = wb.sheets[0]!;
	for (let row = 0; row < 150; row++)
		for (let col = 0; col < 20; col++) edit.setCellInput(0, row, col, `${row}:${col}`);
	for (let row = 0; row < 150; row++)
		sheet.merges.push({ start: { row, col: 0 }, end: { row, col: 1 } });
	const find = vi.spyOn(sheet.merges, 'find');
	const some = vi.spyOn(sheet.merges, 'some');
	try {
		const html = buildPrintHtml(wb, 0, 'Large.xlsx');
		expect(html).toContain('colspan="2"');
		expect(find).not.toHaveBeenCalled();
		expect(some.mock.calls.length).toBeLessThanOrEqual(2);
	} finally {
		find.mockRestore();
		some.mockRestore();
	}
});

it('renders repeated titles and actual numbered header/footer fields on every page', () => {
	const wb = createWorkbook({ sheets: ['Report'] });
	const edit = createEditSession(wb);
	edit.setCellInput(0, 0, 0, 'Repeated title');
	edit.setCellInput(0, 9, 3, 'Last');
	edit.setPageBreak(0, 'row', 5);
	edit.setPageBreak(0, 'col', 2);
	edit.setPrintTitles(0, { rows: { start: 0, end: 0 }, columns: { start: 0, end: 0 } });
	edit.setPageSetup(0, { firstPageNumber: 5, header: '&L&A&RPage &P of &N', footer: '&C&F &D &T' });
	const html = buildPrintHtml(wb, 0, 'Book.xlsx', {
		date: new Date(2026, 9, 12, 13, 4),
		locale: 'en-AU',
	});
	expect(html.match(/Repeated title/g)).toHaveLength(4);
	expect(html).toContain('Page 5 of 4');
	expect(html).toContain('Page 8 of 4');
	expect(html).toContain('Book.xlsx 12/10/26');
	expect(html.match(/xve-print-header/g)).toHaveLength(4);
});
