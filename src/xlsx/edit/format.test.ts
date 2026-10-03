import { describe, expect, it } from 'vitest';
import { parseAddress, parseRange } from '../address.js';
import { getCell } from '../cells.js';
import type { Workbook, Worksheet } from '../model.js';
import { styleAt } from '../styles.js';
import { createWorkbook } from '../workbook.js';
import { mergeWouldDiscard } from './merge.js';
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
const sheet0 = (wb: Workbook): Worksheet => {
	const ws = wb.sheets[0];
	if (!ws) throw new Error('sheet');
	return ws;
};
const style = (wb: Workbook, ref: string) =>
	styleAt(wb, getCell(sheet0(wb), A(ref).row, A(ref).col)?.styleId);
const setup = () => {
	const wb = createWorkbook();
	return { wb, s: createEditSession(wb, { recalc: false }) };
};

describe('applyStyle', () => {
	it('patches every cell of the ranges and shares one style entry', () => {
		const { wb, s } = setup();
		s.applyStyle(0, [R('A1:B2'), R('D4')], { font: { bold: true } });
		expect(style(wb, 'A1').font.bold).toBe(true);
		expect(style(wb, 'B2').font.bold).toBe(true);
		expect(style(wb, 'D4').font.bold).toBe(true);
		expect(wb.styles).toHaveLength(2);
	});
	it('keeps the other properties of each cell', () => {
		const { wb, s } = setup();
		s.applyStyle(0, [R('A1')], { numFmt: '0.00' });
		s.applyStyle(0, [R('A1:A2')], { font: { italic: true } });
		expect(style(wb, 'A1').numFmt).toBe('0.00');
		expect(style(wb, 'A1').font.italic).toBe(true);
		expect(style(wb, 'A2').numFmt).toBe('General');
	});
	it('formats whole columns through the column entry', () => {
		const { wb, s } = setup();
		s.setCellValue(0, 5, 1, 1);
		s.applyStyle(0, [R('B:C')], {
			fill: { type: 'pattern', pattern: 'solid', fgColor: { rgb: 'FFFF00' } },
		});
		const ws = sheet0(wb);
		expect(ws.columns).toHaveLength(1);
		expect(ws.columns[0]).toMatchObject({ min: 1, max: 2 });
		expect(style(wb, 'B6').fill).toMatchObject({ pattern: 'solid' });
		expect(ws.rows.size).toBe(1);
	});
	it('formats whole rows through the row info', () => {
		const { wb, s } = setup();
		s.applyStyle(0, [R('3:3')], { font: { bold: true } });
		const info = sheet0(wb).rowInfo.get(2);
		expect(styleAt(wb, info?.styleId).font.bold).toBe(true);
	});
	it('undoes formatting', () => {
		const { wb, s } = setup();
		s.applyStyle(0, [R('A1:C3')], { alignment: { horizontal: 'center' } });
		s.undo();
		expect(sheet0(wb).rows.size).toBe(0);
		s.applyStyle(0, [R('A:A')], { font: { bold: true } });
		s.undo();
		expect(sheet0(wb).columns).toEqual([]);
	});
});

describe('setBorders', () => {
	it('all draws every edge', () => {
		const { wb, s } = setup();
		s.setBorders(0, R('A1:B2'), 'all');
		for (const ref of ['A1', 'B1', 'A2', 'B2']) {
			const b = style(wb, ref).border;
			expect([b.top?.style, b.bottom?.style, b.left?.style, b.right?.style]).toEqual([
				'thin',
				'thin',
				'thin',
				'thin',
			]);
		}
	});
	it('outside draws only the perimeter', () => {
		const { wb, s } = setup();
		s.setBorders(0, R('A1:C3'), 'outside', { style: 'medium', color: { rgb: 'FF0000' } });
		expect(style(wb, 'A1').border.top).toEqual({ style: 'medium', color: { rgb: 'FF0000' } });
		expect(style(wb, 'A1').border.left?.style).toBe('medium');
		expect(style(wb, 'A1').border.right).toBeUndefined();
		expect(style(wb, 'C3').border.bottom?.style).toBe('medium');
		expect(getCell(sheet0(wb), 1, 1)).toBeUndefined();
	});
	it('thickOutside uses thick edges', () => {
		const { wb, s } = setup();
		s.setBorders(0, R('A1:B1'), 'thickOutside');
		expect(style(wb, 'B1').border.right?.style).toBe('thick');
	});
	it('inside, insideH and insideV draw interior edges only', () => {
		const { wb, s } = setup();
		s.setBorders(0, R('A1:B2'), 'insideH');
		expect(style(wb, 'A1').border.bottom?.style).toBe('thin');
		expect(style(wb, 'A1').border.top).toBeUndefined();
		expect(style(wb, 'A1').border.right).toBeUndefined();
		s.setBorders(0, R('A1:B2'), 'insideV');
		expect(style(wb, 'A1').border.right?.style).toBe('thin');
		expect(style(wb, 'B1').border.left?.style).toBe('thin');
		expect(style(wb, 'B1').border.right).toBeUndefined();
	});
	it('single-side presets touch one edge', () => {
		const { wb, s } = setup();
		s.setBorders(0, R('A1:A3'), 'bottom');
		expect(style(wb, 'A3').border.bottom?.style).toBe('thin');
		expect(getCell(sheet0(wb), 0, 0)).toBeUndefined();
		s.setBorders(0, R('A1:A3'), 'doubleBottom');
		expect(style(wb, 'A3').border.bottom?.style).toBe('double');
	});
	it('none removes every edge', () => {
		const { wb, s } = setup();
		s.setBorders(0, R('A1:B2'), 'all');
		s.setBorders(0, R('A1:B2'), 'none');
		expect(style(wb, 'A1').border).toEqual({});
		expect(sheet0(wb).rows.size).toBe(0);
	});
	it('undoes borders', () => {
		const { wb, s } = setup();
		s.setBorders(0, R('A1:B2'), 'all');
		s.undo();
		expect(sheet0(wb).rows.size).toBe(0);
	});
});

describe('dimensions', () => {
	it('sets column widths and splits spans', () => {
		const { wb, s } = setup();
		sheet0(wb).columns = [{ min: 0, max: 4, width: 12, customWidth: true }];
		s.setColumnWidth(0, [2], 30);
		expect(sheet0(wb).columns).toEqual([
			{ min: 0, max: 1, width: 12, customWidth: true },
			{ min: 2, max: 2, width: 30, customWidth: true },
			{ min: 3, max: 4, width: 12, customWidth: true },
		]);
		s.undo();
		expect(sheet0(wb).columns).toEqual([{ min: 0, max: 4, width: 12, customWidth: true }]);
	});
	it('joins equal neighbouring columns', () => {
		const { wb, s } = setup();
		s.setColumnWidth(0, [0, 1, 2], 20);
		expect(sheet0(wb).columns).toEqual([{ min: 0, max: 2, width: 20, customWidth: true }]);
	});
	it('auto-fits columns with a measure', () => {
		const { wb, s } = setup();
		s.setCellValue(0, 0, 0, 'short');
		s.setCellValue(0, 1, 0, 'a much longer text');
		s.setColumnWidth(0, [0, 1], 'auto', (text) => text.length * 7);
		const col = sheet0(wb).columns.find((c) => c.min === 0);
		expect(col?.width).toBeCloseTo((18 * 7 + 5) / 7, 1);
		expect(col?.bestFit).toBe(true);
		expect(sheet0(wb).columns.find((c) => c.min === 1)).toBeUndefined();
	});
	it('sets and auto-fits row heights', () => {
		const { wb, s } = setup();
		s.setRowHeight(0, [0, 1], 30);
		expect(sheet0(wb).rowInfo.get(1)).toEqual({ height: 30, customHeight: true });
		s.setCellValue(0, 0, 0, 'x');
		s.applyStyle(0, [R('A1')], { font: { size: 22 } });
		s.setRowHeight(0, [0, 1], 'auto');
		expect(sheet0(wb).rowInfo.get(0)?.height).toBe(30);
		expect(sheet0(wb).rowInfo.get(1)).toBeUndefined();
	});
	it('hides and unhides rows and columns', () => {
		const { wb, s } = setup();
		s.setHidden(0, 'row', [2, 3], true);
		s.setHidden(0, 'col', [1], true);
		expect(sheet0(wb).rowInfo.get(3)?.hidden).toBe(true);
		expect(sheet0(wb).columns).toEqual([{ min: 1, max: 1, hidden: true }]);
		s.setHidden(0, 'row', [2], false);
		expect(sheet0(wb).rowInfo.get(2)).toBeUndefined();
		s.setHidden(0, 'col', [1], false);
		expect(sheet0(wb).columns).toEqual([]);
	});
});

describe('merge', () => {
	it('merges, keeping only the top-left value', () => {
		const { wb, s } = setup();
		s.setRangeValues(0, A('A1'), [[1, 2]]);
		s.merge(0, R('A1:B2'), 'merge');
		expect(sheet0(wb).merges).toEqual([R('A1:B2')]);
		expect(getCell(sheet0(wb), 0, 1)).toBeUndefined();
		s.undo();
		expect(getCell(sheet0(wb), 0, 1)?.value).toBe(2);
		expect(sheet0(wb).merges).toEqual([]);
	});
	// Excel 16: B1 =1+1 (0.00) and C1 "c" merged over A1:C1 -> A1 holds =1+1 in 0.00; B3 "r1"
	// and A4 "r2" merged over A3:B4 -> A3 = "r1" (row-major first non-empty).
	it('keeps the first non-empty cell in the corner, with its formula and format', () => {
		const { wb, s } = setup();
		s.setCellInput(0, 0, 1, '=1+1');
		s.applyStyle(0, [R('B1')], { numFmt: '0.00' });
		s.setCellInput(0, 0, 2, 'c');
		expect(mergeWouldDiscard(sheet0(wb), R('A1:C1'))).toBe(true);
		expect(mergeWouldDiscard(sheet0(wb), R('A1:B1'))).toBe(false);
		s.merge(0, R('A1:C1'), 'merge');
		expect(getCell(sheet0(wb), 0, 0)?.formula).toBe('1+1');
		expect(style(wb, 'A1').numFmt).toBe('0.00');
		expect(getCell(sheet0(wb), 0, 1)?.formula).toBeUndefined();
		expect(getCell(sheet0(wb), 0, 2)?.value ?? null).toBeNull();
		s.undo();
		expect(getCell(sheet0(wb), 0, 0)).toBeUndefined();
		expect(getCell(sheet0(wb), 0, 1)?.formula).toBe('1+1');
		s.setCellInput(0, 2, 1, 'r1');
		s.setCellInput(0, 3, 0, 'r2');
		s.merge(0, R('A3:B4'), 'merge');
		expect(getCell(sheet0(wb), 2, 0)?.value).toBe('r1');
		expect(getCell(sheet0(wb), 3, 0)?.value ?? null).toBeNull();
		s.setCellValue(0, 5, 1, 5);
		s.merge(0, R('A6:B6'), 'across');
		expect(getCell(sheet0(wb), 5, 0)?.value).toBe(5);
	});
	it('merge & center centres the result', () => {
		const { wb, s } = setup();
		s.merge(0, R('A1:C1'), 'center');
		expect(style(wb, 'A1').alignment?.horizontal).toBe('center');
	});
	it('merge across merges each row', () => {
		const { wb, s } = setup();
		s.merge(0, R('A1:C3'), 'across');
		expect(sheet0(wb).merges).toEqual([R('A1:C1'), R('A2:C2'), R('A3:C3')]);
	});
	it('replaces overlapping merges and unmerges', () => {
		const { wb, s } = setup();
		s.merge(0, R('A1:B1'), 'merge');
		s.merge(0, R('B1:C2'), 'merge');
		expect(sheet0(wb).merges).toEqual([R('B1:C2')]);
		s.unmerge(0, R('C2'));
		expect(sheet0(wb).merges).toEqual([]);
	});
	it('refuses to merge inside a table', () => {
		const { s } = setup();
		s.setRangeValues(0, A('A1'), [
			['a', 'b'],
			[1, 2],
		]);
		s.createTable(0, R('A1:B2'), true);
		expect(() => s.merge(0, R('A2:B2'), 'merge')).toThrow();
	});
});
