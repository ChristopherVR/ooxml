import { describe, expect, it } from 'vitest';
import { parseAddress, parseRange } from '../address.js';
import { getCell } from '../cells.js';
import { effectiveStyleId } from '../layout/cell-view.js';
import type { Workbook, Worksheet } from '../model.js';
import { styleAt } from '../styles.js';
import { createWorkbook } from '../workbook.js';
import { createEditSession } from './session.js';

const R = (ref: string) => {
	const r = parseRange(ref);
	if (!r) throw new Error(ref);
	return r;
};
const ws = (wb: Workbook): Worksheet => {
	const sheet = wb.sheets[0];
	if (!sheet) throw new Error('sheet');
	return sheet;
};
const shown = (wb: Workbook, ref: string) => {
	const a = parseAddress(ref);
	if (!a) throw new Error(ref);
	return styleAt(wb, effectiveStyleId(ws(wb), a.row, a.col));
};
const stored = (wb: Workbook) =>
	[...ws(wb).rows]
		.flatMap(([row, cells]) =>
			[...cells.keys()].map((col) => String.fromCharCode(65 + col) + (row + 1)),
		)
		.sort();
const yellow = {
	fill: { type: 'pattern', pattern: 'solid', fgColor: { rgb: 'FFFFFF00' } },
} as const;
const red = { fill: { type: 'pattern', pattern: 'solid', fgColor: { rgb: 'FFFF0000' } } } as const;

// The sequence below was run in Excel 16 and saved: Excel shows B3 bold on yellow and E8 italic
// on red, and writes cells B3, D3:F3 (row 3), D5:F5 (row 5) and B8, D8:F8 (row 8) only.
describe('formatting whole rows and columns', () => {
	it('creates the cells where styled rows and columns meet, as Excel does', () => {
		const wb = createWorkbook();
		const s = createEditSession(wb, { recalc: false });
		s.applyStyle(0, [R('B:B')], { font: { bold: true } });
		s.applyStyle(0, [R('3:3')], yellow);
		expect(shown(wb, 'B3').font.bold).toBe(true);
		expect(shown(wb, 'B3').fill).toEqual(yellow.fill);
		s.applyStyle(0, [R('D:F')], { font: { italic: true } });
		s.applyStyle(0, [R('5:5')], { font: { bold: true } });
		expect(stored(wb)).toEqual(['B3', 'D3', 'D5', 'E3', 'E5', 'F3', 'F5']);
		expect(shown(wb, 'E3').font.italic).toBe(true);
		expect(shown(wb, 'E3').fill).toEqual(yellow.fill);
		expect(shown(wb, 'E5').font).toMatchObject({ bold: true, italic: true });
		expect(shown(wb, 'C3').font.bold).toBeFalsy();
		s.applyStyle(0, [R('A:XFD')], { font: { underline: 'single' } });
		expect(shown(wb, 'C3').font.underline).toBe('single');
		expect(shown(wb, 'C3').fill).toEqual(yellow.fill);
		s.applyStyle(0, [R('8:8')], red);
		expect(stored(wb)).toEqual(['B3', 'B8', 'D3', 'D5', 'D8', 'E3', 'E5', 'E8', 'F3', 'F5', 'F8']);
		expect(shown(wb, 'E8').font).toMatchObject({ italic: true, underline: 'single' });
		expect(shown(wb, 'E8').fill).toEqual(red.fill);
		expect(shown(wb, 'C8').font.underline).toBe('single');
		expect(shown(wb, 'C8').fill).toEqual(red.fill);
		expect(shown(wb, 'B8').font).toMatchObject({ bold: true, underline: 'single' });
	});
	it('undoes the intersection cells with the format', () => {
		const wb = createWorkbook();
		const s = createEditSession(wb, { recalc: false });
		s.applyStyle(0, [R('3:3')], yellow);
		s.applyStyle(0, [R('B:B')], { font: { bold: true } });
		expect(getCell(ws(wb), 2, 1)).toBeDefined();
		s.undo();
		expect(getCell(ws(wb), 2, 1)).toBeUndefined();
		expect(shown(wb, 'B3').font.bold).toBeFalsy();
	});
});
