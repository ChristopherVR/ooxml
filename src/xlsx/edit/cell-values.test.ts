import { describe, expect, it } from 'vitest';
import { parseAddress, parseRange } from '../address.js';
import { getCell } from '../cells.js';
import type { Workbook, Worksheet } from '../model.js';
import { applyStylePatch, styleAt } from '../styles.js';
import { createWorkbook } from '../workbook.js';
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
const cell = (wb: Workbook, ref: string) => getCell(sheet0(wb), A(ref).row, A(ref).col);
const fmt = (wb: Workbook, ref: string) => styleAt(wb, cell(wb, ref)?.styleId).numFmt;
const setup = () => {
	const wb = createWorkbook();
	return { wb, s: createEditSession(wb, { recalc: false }) };
};
const type = (text: string) => {
	const { wb, s } = setup();
	s.setCellInput(0, 0, 0, text);
	return { wb, s, c: cell(wb, 'A1') };
};

describe('setCellInput parsing', () => {
	it('stores plain numbers', () => {
		expect(type('42').c?.value).toBe(42);
		expect(type('-3.5').c?.value).toBe(-3.5);
	});
	it('stores text', () => {
		expect(type('hello').c?.value).toBe('hello');
	});
	it('applies a thousands format to grouped numbers', () => {
		const { wb, c } = type('1,234');
		expect(c?.value).toBe(1234);
		expect(fmt(wb, 'A1')).toBe('#,##0');
	});
	it('stores percentages as fractions with a percent format', () => {
		const { wb, c } = type('15%');
		expect(c?.value).toBeCloseTo(0.15);
		expect(fmt(wb, 'A1')).toBe('0%');
	});
	it('recognises currency', () => {
		const { wb, c } = type('$1,234.50');
		expect(c?.value).toBe(1234.5);
		expect(fmt(wb, 'A1')).toContain('$');
	});
	it('recognises dates as serials with a date format', () => {
		const { wb, c } = type('2024-01-15');
		expect(c?.value).toBe(45306);
		expect(fmt(wb, 'A1')).toBe('yyyy-mm-dd');
	});
	it('recognises m/d/yyyy dates', () => {
		expect(type('1/15/2024').c?.value).toBe(45306);
	});
	it('recognises times as day fractions', () => {
		expect(type('12:00').c?.value).toBeCloseTo(0.5);
	});
	it('recognises booleans and errors', () => {
		expect(type('true').c?.value).toBe(true);
		expect(type('FALSE').c?.value).toBe(false);
		expect(type('#N/A').c?.value).toEqual({ error: '#N/A' });
	});
	it('stores formulas without the equals sign', () => {
		const { c } = type('=SUM(A2:A3)');
		expect(c?.formula).toBe('SUM(A2:A3)');
		expect(c?.value).toBeNull();
	});
	it('keeps a leading apostrophe as literal text', () => {
		expect(type("'123").c?.value).toBe('123');
	});
	it('keeps typed text literal in a Text-formatted cell', () => {
		const { wb, s } = setup();
		s.applyStyle(0, [R('A1')], { numFmt: '@' });
		s.setCellInput(0, 0, 0, '=1+1');
		expect(cell(wb, 'A1')?.formula).toBeUndefined();
		expect(cell(wb, 'A1')?.value).toBe('=1+1');
	});
	it('keeps an existing number format', () => {
		const { wb, s } = setup();
		s.applyStyle(0, [R('A1')], { numFmt: '0.000' });
		s.setCellInput(0, 0, 0, '15%');
		expect(fmt(wb, 'A1')).toBe('0.000');
	});
	it('replaces a formula with a constant', () => {
		const { wb, s } = setup();
		s.setCellInput(0, 0, 0, '=1+2');
		s.setCellInput(0, 0, 0, '7');
		expect(cell(wb, 'A1')?.formula).toBeUndefined();
		expect(cell(wb, 'A1')?.value).toBe(7);
	});
	it('empty input clears the value but keeps the format', () => {
		const { wb, s } = setup();
		s.applyStyle(0, [R('A1')], { font: { bold: true } });
		s.setCellInput(0, 0, 0, 'x');
		s.setCellInput(0, 0, 0, '');
		expect(cell(wb, 'A1')?.value).toBeNull();
		expect(styleAt(wb, cell(wb, 'A1')?.styleId).font.bold).toBe(true);
	});
	it('empty input on a plain cell removes it', () => {
		const { wb, s } = setup();
		s.setCellInput(0, 0, 0, 'x');
		s.setCellInput(0, 0, 0, '');
		expect(sheet0(wb).rows.size).toBe(0);
	});
	it('new cells inherit the column format', () => {
		const { wb, s } = setup();
		s.applyStyle(0, [R('C:C')], { font: { italic: true } });
		s.setCellInput(0, 4, 2, 'x');
		expect(styleAt(wb, cell(wb, 'C5')?.styleId).font.italic).toBe(true);
	});
	it('undoes typing including the implied format', () => {
		const { wb, s } = setup();
		s.setCellInput(0, 0, 0, '50%');
		s.undo();
		expect(cell(wb, 'A1')).toBeUndefined();
	});
	it('drops rich text when the value is replaced', () => {
		const { wb, s } = setup();
		sheet0(wb).rows.set(
			0,
			new Map([[0, { value: 'ab', richText: [{ text: 'a' }, { text: 'b' }] }]]),
		);
		s.setCellInput(0, 0, 0, 'c');
		expect(cell(wb, 'A1')?.richText).toBeUndefined();
	});
});

describe('values and clearing', () => {
	it('sets raw values without parsing', () => {
		const { wb, s } = setup();
		s.setCellValue(0, 0, 0, '1,234');
		expect(cell(wb, 'A1')?.value).toBe('1,234');
	});
	it('writes a block of values in one step', () => {
		const { wb, s } = setup();
		s.setRangeValues(0, A('B2'), [
			[1, 2],
			['a', null],
		]);
		expect(cell(wb, 'C2')?.value).toBe(2);
		expect(cell(wb, 'B3')?.value).toBe('a');
		expect(cell(wb, 'C3')).toBeUndefined();
		s.undo();
		expect(sheet0(wb).rows.size).toBe(0);
	});
	it('clears contents but keeps formats', () => {
		const { wb, s } = setup();
		s.setCellValue(0, 0, 0, 1);
		s.applyStyle(0, [R('A1')], { font: { bold: true } });
		s.clearRange(0, R('A1:B2'), 'contents');
		expect(cell(wb, 'A1')?.value).toBeNull();
		expect(cell(wb, 'A1')?.styleId).toBeGreaterThan(0);
	});
	it('clears formats but keeps values', () => {
		const { wb, s } = setup();
		s.setCellValue(0, 0, 0, 1);
		s.applyStyle(0, [R('A1')], { font: { bold: true } });
		s.clearRange(0, R('A1'), 'formats');
		expect(cell(wb, 'A1')?.value).toBe(1);
		expect(cell(wb, 'A1')?.styleId).toBeUndefined();
	});
	it('clear all removes cells, comments, links and merges', () => {
		const { wb, s } = setup();
		s.setCellValue(0, 0, 0, 1);
		s.setComment(0, A('A1'), 'note', 'me');
		s.setHyperlink(0, R('B1'), { target: 'https://example.com' });
		s.merge(0, R('C1:D1'), 'merge');
		s.clearRange(0, R('A1:D1'), 'all');
		const ws = sheet0(wb);
		expect(ws.rows.size).toBe(0);
		expect(ws.comments).toHaveLength(0);
		expect(ws.hyperlinks).toHaveLength(0);
		expect(ws.merges).toHaveLength(0);
		s.undo();
		expect(sheet0(wb).comments).toHaveLength(1);
		expect(sheet0(wb).merges).toHaveLength(1);
	});
	it('clears only comments or only hyperlinks', () => {
		const { wb, s } = setup();
		s.setCellValue(0, 0, 0, 1);
		s.setComment(0, A('A1'), 'note', 'me');
		s.setHyperlink(0, R('A1'), { target: 'https://example.com' });
		s.clearRange(0, R('A1'), 'comments');
		expect(sheet0(wb).comments).toHaveLength(0);
		expect(sheet0(wb).hyperlinks).toHaveLength(1);
		s.clearRange(0, R('A1'), 'hyperlinks');
		expect(sheet0(wb).hyperlinks).toHaveLength(0);
		expect(cell(wb, 'A1')?.value).toBe(1);
	});
	it('clearing a whole column only visits stored cells', () => {
		const { wb, s } = setup();
		s.setCellValue(0, 10, 0, 1);
		s.setCellValue(0, 10, 1, 2);
		s.clearRange(0, R('A:A'), 'contents');
		expect(cell(wb, 'A11')).toBeUndefined();
		expect(cell(wb, 'B11')?.value).toBe(2);
	});
	it('keeps a table column name in step with its header cell', () => {
		const { wb, s } = setup();
		s.setRangeValues(0, A('A1'), [
			['Name', 'Qty'],
			['x', 1],
		]);
		s.createTable(0, R('A1:B2'), true);
		s.setCellInput(0, 0, 1, 'Amount');
		expect(sheet0(wb).tables[0]?.columns[1]?.name).toBe('Amount');
		s.undo();
		expect(sheet0(wb).tables[0]?.columns[1]?.name).toBe('Qty');
	});
	it('uses the inherited row format for new cells', () => {
		const { wb, s } = setup();
		const id = applyStylePatch(wb, 0, { font: { bold: true } });
		sheet0(wb).rowInfo.set(3, { styleId: id });
		s.setCellValue(0, 3, 5, 'x');
		expect(cell(wb, 'F4')?.styleId).toBe(id);
	});
});
