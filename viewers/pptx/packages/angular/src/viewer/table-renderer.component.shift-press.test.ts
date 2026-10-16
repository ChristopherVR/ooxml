/**
 * table-renderer.component.shift-press.test.ts: a Shift press inside the
 * selected table's cell range must stop at the cell.
 *
 * The slide canvas selects on `pointerdown` and its Shift branch toggles the
 * pressed element, so a Shift-click meant to stretch the cell range first took
 * the table out of the selection: the Table Layout tab vanished and Merge Cells
 * could never be reached. Tested through the pure helper, for the same TestBed
 * reason as `table-renderer.component.drilldown.test.ts`.
 */
import { describe, expect, it } from 'vitest';

import { cellPressConsumesShift } from './table-renderer.component';

const BASE = {
	canDrillDown: true,
	shiftKey: true,
	elementSelected: true,
	rangeElementId: 'tbl1',
	elementId: 'tbl1',
};

describe('cellPressConsumesShift', () => {
	it('consumes a Shift press that extends the range on the selected table', () => {
		expect(cellPressConsumesShift(BASE)).toBeTruthy();
	});

	it('lets a plain press through to the canvas', () => {
		expect(cellPressConsumesShift({ ...BASE, shiftKey: false })).toBeFalsy();
	});

	it('lets a Shift press through when the table is not selected', () => {
		expect(cellPressConsumesShift({ ...BASE, elementSelected: false })).toBeFalsy();
	});

	it('lets a Shift press through when the range belongs to another table', () => {
		expect(cellPressConsumesShift({ ...BASE, rangeElementId: 'tbl2' })).toBeFalsy();
		expect(cellPressConsumesShift({ ...BASE, rangeElementId: undefined })).toBeFalsy();
	});

	it('never consumes when the table cannot be drilled into', () => {
		expect(cellPressConsumesShift({ ...BASE, canDrillDown: false })).toBeFalsy();
	});
});
