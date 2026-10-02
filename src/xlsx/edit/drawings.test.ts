import { readFileSync } from 'node:fs';
import path from 'node:path';
import { describe, expect, it } from 'vitest';
import type { DrawingAnchor } from '../model.js';
import { loadXlsx } from '../read/index.js';
import { saveXlsx } from '../write/index.js';
import { deleteDrawing, setDrawingAnchor } from './drawings.js';
import { testContext } from './test-context.js';

const fixture = () =>
	new Uint8Array(
		readFileSync(path.join(import.meta.dirname, '..', '__fixtures__', 'excel-features.xlsx')),
	);

const moved: DrawingAnchor = {
	from: { row: 20, col: 6, rowOffset: 0, colOffset: 9525 },
	to: { row: 30, col: 12, rowOffset: 19050, colOffset: 0 },
};

describe('drawing commands', () => {
	it('moves a drawing as one undoable step and saves the new anchor', async () => {
		const wb = await loadXlsx(fixture());
		const s = wb.sheets.findIndex((sheet) => sheet.drawings.length > 0);
		expect(s).toBeGreaterThanOrEqual(0);
		const ctx = testContext(wb);
		const before = structuredClone(wb.sheets[s]?.drawings[0]?.anchor);
		setDrawingAnchor(ctx, s, 0, moved);
		expect(wb.sheets[s]?.drawings[0]?.anchor).toEqual(moved);
		expect(ctx.steps).toHaveLength(1);

		const reread = await loadXlsx(await saveXlsx(wb));
		expect(reread.sheets[s]?.drawings[0]?.anchor.from).toMatchObject(moved.from);
		expect(reread.sheets[s]?.drawings[0]?.anchor.to).toMatchObject(moved.to ?? {});

		ctx.undo();
		expect(wb.sheets[s]?.drawings[0]?.anchor).toEqual(before);
	});

	it('deletes a drawing and the saved file no longer has it', async () => {
		const wb = await loadXlsx(fixture());
		const s = wb.sheets.findIndex((sheet) => sheet.drawings.length > 0);
		const count = wb.sheets[s]?.drawings.length ?? 0;
		const ctx = testContext(wb);
		deleteDrawing(ctx, s, 0);
		expect(wb.sheets[s]?.drawings).toHaveLength(count - 1);
		const reread = await loadXlsx(await saveXlsx(wb));
		expect(reread.sheets[s]?.drawings).toHaveLength(count - 1);
		ctx.undo();
		expect(wb.sheets[s]?.drawings).toHaveLength(count);
	});

	it('rejects bad indices and anchors', async () => {
		const wb = await loadXlsx(fixture());
		const s = wb.sheets.findIndex((sheet) => sheet.drawings.length > 0);
		const ctx = testContext(wb);
		expect(() => setDrawingAnchor(ctx, s, 99, moved)).toThrow(RangeError);
		expect(() => deleteDrawing(ctx, s, 99)).toThrow(RangeError);
		expect(() =>
			setDrawingAnchor(ctx, s, 0, { from: { row: -1, col: 0, rowOffset: 0, colOffset: 0 } }),
		).toThrow(RangeError);
		expect(ctx.steps).toHaveLength(0);
	});
});
