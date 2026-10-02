import { createWorksheet, MAX_COL, MAX_ROW } from '@christophervr/xlsx-core';
import { describe, expect, it } from 'vitest';
import type { Selection } from '../context.js';
import {
	addRange,
	colSelected,
	expandForMerges,
	extendSelection,
	moveWithinSelection,
	movingEnd,
	rangeRef,
	rowSelected,
	rowsRange,
	colsRange,
	selectCell,
	selectionRef,
	spanRange,
} from './selection-ops.js';

const sel = (ranges: Selection['ranges'], active = ranges[0]!.start): Selection => ({
	sheet: 0,
	active,
	anchor: active,
	ranges,
});

describe('selection ops', () => {
	it('expands a range over merges it touches', () => {
		const sheet = createWorksheet('S', 1);
		sheet.merges.push({ start: { row: 1, col: 1 }, end: { row: 3, col: 2 } });
		expect(expandForMerges(sheet, spanRange({ row: 0, col: 0 }, { row: 1, col: 1 }))).toEqual({
			start: { row: 0, col: 0 },
			end: { row: 3, col: 2 },
		});
		expect(selectCell(0, sheet, { row: 2, col: 2 }).ranges[0]).toEqual({
			start: { row: 1, col: 1 },
			end: { row: 3, col: 2 },
		});
	});

	it('extends from the anchor and tracks the moving corner', () => {
		const start = selectCell(0, undefined, { row: 4, col: 4 });
		const up = extendSelection(start, undefined, { row: 2, col: 6 });
		expect(up.ranges[0]).toEqual({ start: { row: 2, col: 4 }, end: { row: 4, col: 6 } });
		expect(up.active).toEqual({ row: 4, col: 4 });
		expect(movingEnd(up)).toEqual({ row: 2, col: 6 });
	});

	it('adds ranges with Ctrl and formats the reference', () => {
		const two = addRange(selectCell(0, undefined, { row: 0, col: 0 }), undefined, {
			row: 3,
			col: 2,
		});
		expect(two.ranges).toHaveLength(2);
		expect(two.active).toEqual({ row: 3, col: 2 });
		expect(selectionRef(two)).toBe('A1,C4');
		expect(rangeRef(rowsRange(1, 3))).toBe('2:4');
		expect(rangeRef(colsRange(0, 2))).toBe('A:C');
		expect(rangeRef(colsRange(3, 3))).toBe('D:D');
	});

	it('walks a selection with Enter (down) and Tab (right), wrapping', () => {
		const block = sel([{ start: { row: 0, col: 0 }, end: { row: 1, col: 1 } }]);
		const a = moveWithinSelection(block, 'down')!;
		expect(a.active).toEqual({ row: 1, col: 0 });
		const b = moveWithinSelection(a, 'down')!;
		expect(b.active).toEqual({ row: 0, col: 1 });
		const c = moveWithinSelection(block, 'right')!;
		expect(c.active).toEqual({ row: 0, col: 1 });
		const last = moveWithinSelection({ ...block, active: { row: 1, col: 1 } }, 'down')!;
		expect(last.active).toEqual({ row: 0, col: 0 });
		expect(
			moveWithinSelection(sel([{ start: { row: 0, col: 0 }, end: { row: 0, col: 0 } }]), 'down'),
		).toBeUndefined();
	});

	it('reports header highlight states', () => {
		const rows = sel([rowsRange(2, 3)]);
		expect(rowSelected(rows, 2)).toBe('full');
		expect(rowSelected(rows, 4)).toBe('none');
		expect(colSelected(rows, 5)).toBe('part');
		const cols = sel([{ start: { row: 0, col: 1 }, end: { row: MAX_ROW, col: 1 } }]);
		expect(colSelected(cols, 1)).toBe('full');
		expect(rowSelected(cols, 7)).toBe('part');
		expect(colsRange(0, MAX_COL).end.col).toBe(MAX_COL);
	});
});
