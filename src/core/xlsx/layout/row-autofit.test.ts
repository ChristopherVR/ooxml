import { describe, expect, it } from 'vitest';
import { putCell } from '../cells.js';
import { internStyle } from '../styles.js';
import { createWorkbook, defaultCellStyle } from '../workbook.js';
import {
	approximateMeasure,
	autoFitRowHeight,
	lineHeightPoints,
	wrappedLineCount,
} from './row-autofit.js';

const measure = (text: string): number => text.length * 7;

describe('lineHeightPoints', () => {
	it('matches Excel row heights for Calibri sizes', () => {
		expect(lineHeightPoints(11)).toBe(15);
		expect(lineHeightPoints(10)).toBe(12.75);
		expect(lineHeightPoints(18)).toBe(23.25);
	});

	it('interpolates other sizes to whole pixels', () => {
		const h = lineHeightPoints(13);
		expect(h).toBeGreaterThan(15.75);
		expect(h).toBeLessThan(18.75);
		expect((h / 0.75) % 1).toBe(0);
	});
});

describe('wrappedLineCount', () => {
	it('wraps words and keeps explicit breaks', () => {
		expect(wrappedLineCount('one two three', 70, measure)).toBe(2);
		expect(wrappedLineCount('a\nb\nc', 100, measure)).toBe(3);
		expect(wrappedLineCount('short', 100, measure)).toBe(1);
	});

	it('breaks a word longer than the cell', () => {
		expect(wrappedLineCount('abcdefghij', 35, measure)).toBe(2);
	});
});

describe('autoFitRowHeight', () => {
	it('uses the default height for empty rows and plain text', () => {
		const wb = createWorkbook();
		const sheet = wb.sheets[0];
		if (!sheet) throw new Error('no sheet');
		expect(autoFitRowHeight(wb, 0, 0, approximateMeasure)).toBe(sheet.defaultRowHeight);
		putCell(sheet, 0, 0, { value: 'hello' });
		expect(autoFitRowHeight(wb, 0, 0, approximateMeasure)).toBe(15);
	});

	it('grows for larger fonts and wrapped text', () => {
		const wb = createWorkbook();
		const sheet = wb.sheets[0];
		if (!sheet) throw new Error('no sheet');
		const base = defaultCellStyle();
		const big = internStyle(wb, { ...base, font: { ...base.font, size: 18 } });
		const wrap = internStyle(wb, { ...base, alignment: { wrapText: true } });
		putCell(sheet, 0, 0, { value: 'big', styleId: big });
		expect(autoFitRowHeight(wb, 0, 0, approximateMeasure)).toBe(23.25);
		putCell(sheet, 1, 0, { value: 'word '.repeat(20).trim(), styleId: wrap });
		expect(autoFitRowHeight(wb, 0, 1, (t) => measure(t))).toBeGreaterThanOrEqual(45);
		putCell(sheet, 2, 0, { value: 'word '.repeat(20).trim() });
		expect(autoFitRowHeight(wb, 0, 2, (t) => measure(t))).toBe(15);
	});

	it('ignores merged cells', () => {
		const wb = createWorkbook();
		const sheet = wb.sheets[0];
		if (!sheet) throw new Error('no sheet');
		const base = defaultCellStyle();
		const big = internStyle(wb, { ...base, font: { ...base.font, size: 28 } });
		putCell(sheet, 0, 0, { value: 'big', styleId: big });
		sheet.merges.push({ start: { row: 0, col: 0 }, end: { row: 1, col: 1 } });
		expect(autoFitRowHeight(wb, 0, 0, approximateMeasure)).toBe(sheet.defaultRowHeight);
	});
});
