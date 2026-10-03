import { describe, expect, it } from 'vitest';
import { getCell } from '../cells.js';
import { isSpilledCell } from './spill.js';
import { book, E, engine, get, grid, locate, set } from './test-helpers.js';

describe('reusing spill destinations', () => {
	it('updates same-size spills and their readers while retaining children and formatting', () => {
		// Insert the reader first to require ordering through the live spill on the second pass.
		const wb = book({ D1: '=B3*10', C1: 1, A1: '=SEQUENCE(3,2,C1)' });
		const ws = wb.sheets[0]!;
		const e = engine(wb);
		e.recalculateAll();
		const child = getCell(ws, 2, 1)!;
		child.styleId = 1;
		const marker = isSpilledCell(child) ? child.spillAnchor : undefined;
		set(wb, 'C1', 10);
		e.recalculateAll();
		expect(grid(wb, 'A1:B3')).toEqual([
			[10, 11],
			[12, 13],
			[14, 15],
		]);
		expect(get(wb, 'D1')).toBe(150);
		expect(getCell(ws, 2, 1)).toBe(child);
		expect(child.styleId).toBe(1);
		expect(isSpilledCell(child) && child.spillAnchor).toBe(marker);
		set(wb, 'C1', 20);
		e.recalculateFrom([locate(wb, 'C1')]);
		expect(get(wb, 'D1')).toBe(250);
		expect(getCell(ws, 2, 1)).toBe(child);
	});

	it('rechecks newly merged destinations and clears the old spill before recovering', () => {
		const wb = book({ A1: '=SEQUENCE(3)' });
		const ws = wb.sheets[0]!;
		const e = engine(wb);
		e.recalculateAll();
		ws.merges.push({ start: { row: 1, col: 0 }, end: { row: 1, col: 1 } });
		e.recalculateAll();
		expect(get(wb, 'A1')).toEqual(E.SPILL);
		expect(grid(wb, 'A2:A3')).toEqual([[null], [null]]);
		expect(e.spillRange(0, 0, 0)).toBeUndefined();
		ws.merges.length = 0;
		e.recalculateAll();
		expect(grid(wb, 'A1:A3')).toEqual([[1], [2], [3]]);
	});

	it('preserves an overwrite but releases the other old children when blocked', () => {
		const wb = book({ A1: '=SEQUENCE(3)' });
		const e = engine(wb);
		e.recalculateAll();
		set(wb, 'A2', 'mine');
		e.recalculateAll();
		expect(grid(wb, 'A1:A3')).toEqual([[E.SPILL], ['mine'], [null]]);
	});

	it('rejects a formula inserted into a destination even with its old ownership marker', () => {
		const wb = book({ A1: '=SEQUENCE(3)' });
		const e = engine(wb);
		e.recalculateAll();
		const child = getCell(wb.sheets[0]!, 1, 0)!;
		child.formula = '99';
		e.recalculateFrom([locate(wb, 'A2')]);
		expect(grid(wb, 'A1:A3')).toEqual([[E.SPILL], [99], [null]]);
		expect(child.formula).toBe('99');
	});
});
