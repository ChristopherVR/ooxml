import {
	createWorkbook,
	putCell,
	cellView,
	createGridMetrics,
	internStyle,
	defaultCellStyle,
} from '@christophervr/xlsx-core';
import { describe, expect, it } from 'vitest';
import { buildItems } from './cell-items.js';
import { cssFont, edgeCss, fillPaint, hashes, patternImage } from './cell-paint.js';
import { iconSvg } from './icon-sets.js';

describe('cell paint', () => {
	it('builds font shorthands with fallbacks and zoom', () => {
		const font = {
			family: 'Calibri',
			sizePx: 14.67,
			bold: true,
			italic: true,
			strike: false,
			color: '#000000',
		};
		expect(cssFont(font, 200)).toBe(
			'italic bold 29.34px "Calibri", "Carlito", "Segoe UI", Arial, sans-serif',
		);
		expect(cssFont({ ...font, family: 'My "Font"', bold: false, italic: false })).toContain(
			'"My Font"',
		);
	});

	it('turns fills and borders into CSS', () => {
		expect(fillPaint({ background: '#FF0000' })).toEqual({ background: '#FF0000' });
		expect(fillPaint({ pattern: 'solid', fg: '#00FF00', bg: '#FFFFFF' })).toEqual({
			background: '#00FF00',
		});
		const pattern = fillPaint({ pattern: 'lightGrid', fg: '#123456', bg: '#FFFFFF' });
		expect(pattern?.background).toBe('#FFFFFF');
		expect(pattern?.image).toContain('data:image/svg+xml');
		expect(patternImage('none', '#000')).toBeUndefined();
		expect(edgeCss({ widthPx: 1, style: 'double', color: '#000' })).toBe('3px double #000');
		expect(hashes(30, 7)).toBe('####');
		expect(iconSvg('3Arrows', 0)).toContain('<svg');
	});

	it('lets text overflow into empty neighbours and shows #### for numbers that do not fit', () => {
		const wb = createWorkbook();
		const sheet = wb.sheets[0]!;
		putCell(sheet, 0, 0, { value: 'A long piece of text that spills' });
		putCell(sheet, 0, 3, { value: 'stop' });
		const narrow = internStyle(wb, { ...defaultCellStyle(), numFmt: '0.00' });
		putCell(sheet, 1, 0, { value: 123456789012, styleId: narrow });
		const metrics = createGridMetrics(sheet, { zoom: 100 });
		const items = buildItems({
			sheet,
			metrics,
			zoom: 100,
			rows: [0, 1],
			cols: [0, 1, 2, 3],
			view: (r, c) => (sheet.rows.get(r)?.get(c) ? cellView(wb, 0, r, c) : null),
			measure: (text) => text.length * 7,
			hasFormula: () => false,
		});
		const a1 = items.find((i) => i.key === '0:0')!;
		expect(a1.overflowing).toBe(true);
		expect(a1.textWidth).toBe(metrics.colLeft(3));
		const a2 = items.find((i) => i.key === '1:0')!;
		expect(a2.text).toMatch(/^#+$/);
	});

	it('paints merged anchors at the merged size and skips covered cells', () => {
		const wb = createWorkbook();
		const sheet = wb.sheets[0]!;
		putCell(sheet, 0, 0, { value: 'Merged' });
		sheet.merges.push({ start: { row: 0, col: 0 }, end: { row: 1, col: 2 } });
		const metrics = createGridMetrics(sheet, { zoom: 100 });
		const items = buildItems({
			sheet,
			metrics,
			zoom: 100,
			rows: [1, 2],
			cols: [1, 2],
			view: (r, c) => (sheet.rows.get(r)?.get(c) ? cellView(wb, 0, r, c) : null),
			measure: () => 10,
			hasFormula: () => false,
		});
		expect(items).toHaveLength(1);
		expect(items[0]?.w).toBe(metrics.colLeft(3));
		expect(items[0]?.merged).toBe(true);
	});
});
