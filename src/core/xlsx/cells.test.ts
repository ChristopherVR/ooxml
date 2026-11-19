import { describe, expect, it } from 'vitest';
import type { CellRange } from './address';
import { deleteCell, forEachCellInRange, getCell, putCell } from './cells';
import { createWorkbook } from './workbook';

const sheet = () => createWorkbook().sheets[0]!;
const range = (r1: number, c1: number, r2: number, c2: number): CellRange => ({
	start: { row: r1, col: c1 },
	end: { row: r2, col: c2 },
});
const visited = (ws: ReturnType<typeof sheet>, r: CellRange) => {
	const out: [number, number][] = [];
	forEachCellInRange(ws, r, (_cell, row, col) => out.push([row, col]));
	return out;
};

describe('forEachCellInRange', () => {
	it('visits stored cells in row-major order however they were inserted', () => {
		const ws = sheet();
		for (const [row, col] of [
			[9, 3],
			[2, 8],
			[2, 1],
			[500, 0],
			[9, 0],
		] as const)
			putCell(ws, row, col, { value: 1 });
		const expected = [
			[2, 1],
			[2, 8],
			[9, 0],
			[9, 3],
			[500, 0],
		];
		// A span narrower than the sheet, one wider, and whole columns and rows.
		expect(visited(ws, range(0, 0, 600, 9))).toEqual(expected);
		expect(visited(ws, range(0, 0, 1_048_575, 16_383))).toEqual(expected);
		expect(visited(ws, range(2, 1, 9, 3))).toEqual([
			[2, 1],
			[9, 3],
		]);
		expect(visited(ws, range(3, 0, 8, 9))).toEqual([]);
	});

	it('lets the visitor delete and replace the cell it is given', () => {
		const ws = sheet();
		for (let row = 0; row < 5; row++)
			for (let col = 0; col < 5; col++) putCell(ws, row, col, { value: row * col });
		forEachCellInRange(ws, range(1, 1, 3, 3), (cell, row, col) => {
			if (row === col) deleteCell(ws, row, col);
			else putCell(ws, row, col, { value: (cell.value as number) + 100 });
		});
		expect(getCell(ws, 2, 2)).toBeUndefined();
		expect(getCell(ws, 1, 2)?.value).toBe(102);
		expect(getCell(ws, 0, 0)?.value).toBe(0);
	});

	it('costs the size of a small range, not of the sheet', () => {
		const ws = sheet();
		for (let row = 0; row < 300_000; row++) putCell(ws, row, 0, { value: row });
		const start = performance.now();
		let seen = 0;
		for (let i = 0; i < 20_000; i++)
			forEachCellInRange(ws, range(i * 10, 0, i * 10, 0), () => seen++);
		expect(seen).toBe(20_000);
		expect(performance.now() - start).toBeLessThan(process.env.CI ? 600 : 200);
	});
});
