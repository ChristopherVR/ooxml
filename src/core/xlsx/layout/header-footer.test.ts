import { expect, it } from 'vitest';
import { headerFooterImageKey, renderHeaderFooterRuns, parseHeaderFooter } from './header-footer';
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
		'&"Arial,Unknown"',
	]);
});

// Ground truth: Excel 16 exported "[&P+2]" as 3, 4, 5 on pages 1..3, and these colours for
// "&K05+000", "&K05+040", "&K05-040", "&K00+000" and "&K03+000" in the Office 2023 theme
// (E97132, F1A983, 983F10, FFFFFF, 0E2841).
it('applies page-number offsets like Excel', () => {
	const at = (page: number, code: string) =>
		renderHeaderFooterRuns(code, { ...context, page })
			.sections.center.map((r) => r.text)
			.join('');
	expect(at(1, '[&P+2]')).toBe('[3]');
	expect(at(3, '[&P-1]')).toBe('[2]');
	expect(at(1, '[&P-1]')).toBe('[0]');
	expect(at(1, '[&P-7]')).toBe('[17]');
});

it('resolves theme colours, double underline, super and subscript', () => {
	const themeColors = ['FFFFFF', '000000', 'E8E8E8', '0E2841', '156082', 'E97132'];
	const run = (code: string) =>
		renderHeaderFooterRuns(code, { ...context, themeColors }).sections.center[0]!.style;
	expect(run('&K05+000x').color).toBe('#E97132');
	// Excel's tinted colours differ from this tint arithmetic by at most 2 per channel.
	const near = (css: string | undefined, expected: string) =>
		[1, 3, 5].every(
			(i) =>
				Math.abs(
					Number.parseInt((css ?? '#000000').slice(i, i + 2), 16) -
						Number.parseInt(expected.slice(i, i + 2), 16),
				) <= 2,
		);
	expect(near(run('&K05+040x').color, '#F1A983')).toBe(true);
	expect(near(run('&K05-040x').color, '#983F10')).toBe(true);
	expect(run('&K00+000x').color).toBe('#FFFFFF');
	expect(run('&K03+000x').color).toBe('#0E2841');
	// A later RGB colour replaces a theme colour.
	expect(run('&K05+000&KFF0000x')).toEqual({ color: '#FF0000' });
	expect(run('&Edouble').doubleUnderline).toBe(true);
	expect(run('&Xup').superscript).toBe(true);
	expect(run('&Ydown').subscript).toBe(true);
});

it('marks the pictures of a header section', () => {
	const runs = renderHeaderFooterRuns('&L&G&CTitle&R&G', context).sections;
	expect(runs.left).toEqual([{ text: '', style: {}, picture: true }]);
	expect(runs.right[0]?.picture).toBe(true);
	expect(headerFooterImageKey('header', 'left')).toBe('LH');
	expect(headerFooterImageKey('footer', 'right', 'even')).toBe('RFEVEN');
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
