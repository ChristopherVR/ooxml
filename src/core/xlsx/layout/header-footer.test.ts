import { expect, it } from 'vitest';
import { renderHeaderFooterRuns, parseHeaderFooter } from './header-footer';
import { createWorkbook } from '../workbook';
import { createEditSession } from '../edit/session';
import { saveXlsx } from '../write/index';
import { loadXlsx } from '../read/index';

const context = { page: 3, pages: 4, sheet: 'Report', file: 'Book.xlsx' };
it('applies native font and RGB controls to fields, toggles styles and isolates sections', () => {
	const result = renderHeaderFooterRuns(
		'&L&BLeft&C&"Arial,Bold Italic"&18&KFF0000Page &P&B plain&RRight',
		context,
	);
	expect(result.sections.left[0]).toEqual({ text: 'Left', style: { bold: true } });
	expect(result.sections.center[0]).toEqual({
		text: 'Page 3',
		style: { font: 'Arial', bold: true, italic: true, size: 18, color: '#FF0000' },
	});
	expect(result.sections.center.at(-1)?.style.bold).toBe(false);
	expect(result.sections.right[0]).toEqual({ text: 'Right', style: {} });
	expect(result.unsupported).toEqual([]);
	expect(parseHeaderFooter('&K01+050&G&P+2&E&"Arial,Unknown"').unsupported).toEqual([
		'&K01+050',
		'&G',
		'&P+2',
		'&E',
		'&"Arial,Unknown"',
	]);
});

it('preserves rich codes and unsupported controls through edit, undo and save/reopen', async () => {
	const wb = createWorkbook();
	const edit = createEditSession(wb);
	const header = '&C&"Arial,Bold"&18&KFF0000Report &P&U under&U &G';
	edit.setPageSetup(0, { header });
	edit.setPageSetup(0, { header: '&CChanged' });
	edit.undo();
	expect((await loadXlsx(await saveXlsx(wb))).sheets[0]!.pageSetup!.header).toBe(header);
});
