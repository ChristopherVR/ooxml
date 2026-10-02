import { describe, expect, it } from 'vitest';
import { MAX_ROW } from '../address.js';
import { putCell } from '../cells.js';
import type { CellStyle } from '../model.js';
import { internStyle } from '../styles.js';
import { createWorkbook, defaultCellStyle } from '../workbook.js';
import { autoFitColumnWidth } from './autofit.js';
import { createGridMetrics } from './metrics.js';
import { isOverflowTarget, mergeView, overflowExtent, selectionStats } from './sheet-queries.js';
import { columnWidthToPixels } from './units.js';

const range = (r1: number, c1: number, r2: number, c2: number) => ({
	start: { row: r1, col: c1 },
	end: { row: r2, col: c2 },
});

describe('mergeView', () => {
	const wb = createWorkbook();
	const sheet = wb.sheets[0]!;
	sheet.merges.push(range(1, 1, 2, 3));

	it('marks the anchor and the covered cells', () => {
		expect(mergeView(sheet, 1, 1)).toEqual({
			anchor: true,
			hidden: false,
			range: range(1, 1, 2, 3),
		});
		expect(mergeView(sheet, 2, 3)).toEqual({
			anchor: false,
			hidden: true,
			range: range(1, 1, 2, 3),
		});
		expect(mergeView(sheet, 0, 0)).toEqual({ anchor: false, hidden: false });
	});
});

describe('selectionStats', () => {
	const wb = createWorkbook();
	const sheet = wb.sheets[0]!;
	putCell(sheet, 0, 0, { value: 1 });
	putCell(sheet, 1, 0, { value: 2 });
	putCell(sheet, 2, 0, { value: 'text' });
	putCell(sheet, 3, 0, { value: true });
	putCell(sheet, 4, 0, { value: 6 });
	putCell(sheet, 5, 0, { value: null, styleId: 0 });
	putCell(sheet, 0, 1, { value: -3 });

	it('counts, sums and averages like the status bar', () => {
		expect(selectionStats(wb, 0, [range(0, 0, 5, 0)])).toEqual({
			count: 5,
			numericCount: 3,
			sum: 9,
			average: 3,
			min: 1,
			max: 6,
		});
	});

	it('handles whole columns and overlapping ranges', () => {
		const stats = selectionStats(wb, 0, [range(0, 0, MAX_ROW, 1), range(0, 0, 1, 0)]);
		expect(stats.count).toBe(6);
		expect(stats.sum).toBe(6);
		expect(stats.min).toBe(-3);
	});

	it('omits average, min and max without numbers', () => {
		expect(selectionStats(wb, 0, [range(2, 0, 3, 0)])).toEqual({
			count: 2,
			numericCount: 0,
			sum: 0,
		});
		expect(selectionStats(wb, 9, [range(0, 0, 1, 1)]).count).toBe(0);
		expect(selectionStats(wb, sheet, [range(0, 0, 0, 0)]).sum).toBe(1);
	});
});

describe('overflowExtent', () => {
	const wb = createWorkbook();
	const sheet = wb.sheets[0]!;
	putCell(sheet, 0, 2, { value: 'long text here' });
	putCell(sheet, 0, 5, { value: 'blocker' });
	putCell(sheet, 0, 1, { value: '' });
	const metrics = createGridMetrics(sheet);

	it('keeps text that fits inside its cell', () => {
		expect(overflowExtent(sheet, metrics, 0, 2, 50, 'left')).toMatchObject({
			startCol: 2,
			endCol: 2,
			left: 128,
			right: 192,
		});
	});

	it('spills left-aligned text to the right until it fits', () => {
		expect(overflowExtent(sheet, metrics, 0, 2, 100, 'left')).toMatchObject({
			startCol: 2,
			endCol: 3,
			right: 256,
		});
	});

	it('stops at the first non-empty cell', () => {
		expect(overflowExtent(sheet, metrics, 0, 2, 1000, 'left').endCol).toBe(4);
	});

	it('spills right-aligned text to the left over empty cells', () => {
		expect(overflowExtent(sheet, metrics, 0, 2, 200, 'right')).toMatchObject({
			startCol: 0,
			endCol: 2,
			left: 0,
		});
	});

	it('spills centred text both ways', () => {
		expect(overflowExtent(sheet, metrics, 0, 2, 180, 'center')).toMatchObject({
			startCol: 1,
			endCol: 3,
		});
	});

	it('does not spill into merges', () => {
		sheet.merges.push(range(0, 3, 0, 4));
		expect(isOverflowTarget(sheet, 0, 3)).toBe(false);
		expect(overflowExtent(sheet, metrics, 0, 2, 1000, 'left').endCol).toBe(2);
		sheet.merges.pop();
	});
});

describe('autoFitColumnWidth', () => {
	const measure = (text: string): number => text.length * 7;

	it('fits the widest display text plus padding', () => {
		const wb = createWorkbook();
		const sheet = wb.sheets[0]!;
		putCell(sheet, 0, 0, { value: 'abc' });
		putCell(sheet, 1, 0, { value: 'abcdefghij' });
		putCell(sheet, 2, 0, { value: 12345 });
		const width = autoFitColumnWidth(sheet, wb, 0, measure);
		expect(columnWidthToPixels(width)).toBe(70 + 6);
	});

	it('uses the longest line, indent and rich runs', () => {
		const wb = createWorkbook();
		const sheet = wb.sheets[0]!;
		const style: CellStyle = { ...defaultCellStyle(), alignment: { indent: 1, wrapText: true } };
		putCell(sheet, 0, 0, { value: 'ab\nabcdef', styleId: internStyle(wb, style) });
		putCell(sheet, 1, 0, { value: 'xy', richText: [{ text: 'x' }, { text: 'y' }] });
		expect(columnWidthToPixels(autoFitColumnWidth(0, wb, 0, measure))).toBe(42 + 9 + 6);
	});

	it('ignores merged cells and keeps the default width for empty columns', () => {
		const wb = createWorkbook();
		const sheet = wb.sheets[0]!;
		putCell(sheet, 0, 0, { value: 'a very long merged title' });
		sheet.merges.push(range(0, 0, 0, 3));
		expect(columnWidthToPixels(autoFitColumnWidth(sheet, wb, 0, measure))).toBe(64);
		sheet.defaultColWidth = 13;
		expect(autoFitColumnWidth(sheet, wb, 7, measure)).toBe(13);
	});
});
