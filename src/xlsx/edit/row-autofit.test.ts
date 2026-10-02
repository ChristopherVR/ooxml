import { describe, expect, it } from 'vitest';
import { putCell } from '../cells.js';
import { internStyle } from '../styles.js';
import { createWorkbook, defaultCellStyle } from '../workbook.js';
import { autoGrowRows } from './row-autofit.js';
import { testContext } from './test-context.js';

function setup() {
	const wb = createWorkbook();
	const sheet = wb.sheets[0];
	if (!sheet) throw new Error('no sheet');
	const base = defaultCellStyle();
	const big = internStyle(wb, { ...base, font: { ...base.font, size: 18 } });
	return { wb, sheet, big };
}

describe('autoGrowRows', () => {
	it('grows a row without a custom height, undoably', () => {
		const { wb, sheet, big } = setup();
		putCell(sheet, 2, 0, { value: 'Title', styleId: big });
		const ctx = testContext(wb);
		expect(autoGrowRows(ctx, 0, [2])).toEqual([2]);
		expect(sheet.rowInfo.get(2)?.height).toBe(23.25);
		expect(sheet.rowInfo.get(2)?.customHeight).toBeUndefined();
		ctx.undo();
		expect(sheet.rowInfo.get(2)).toBeUndefined();
	});

	it('shrinks back to the default and records nothing when unchanged', () => {
		const { wb, sheet } = setup();
		sheet.rowInfo.set(1, { height: 30 });
		putCell(sheet, 1, 0, { value: 'plain' });
		const ctx = testContext(wb);
		expect(autoGrowRows(ctx, 0, [1])).toEqual([1]);
		expect(sheet.rowInfo.has(1)).toBe(false);
		expect(autoGrowRows(ctx, 0, [1, 5])).toEqual([]);
		expect(ctx.steps).toHaveLength(1);
	});

	it('leaves custom-height and hidden rows alone', () => {
		const { wb, sheet, big } = setup();
		putCell(sheet, 0, 0, { value: 'x', styleId: big });
		putCell(sheet, 1, 0, { value: 'x', styleId: big });
		sheet.rowInfo.set(0, { height: 12, customHeight: true });
		sheet.rowInfo.set(1, { hidden: true });
		const ctx = testContext(wb);
		expect(autoGrowRows(ctx, 0, [0, 1])).toEqual([]);
		expect(sheet.rowInfo.get(0)?.height).toBe(12);
	});
});
