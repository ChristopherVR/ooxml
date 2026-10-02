// @vitest-environment jsdom
import {
	createGridMetrics,
	createWorkbook,
	putCell,
	type ChartObject,
} from '@christophervr/xlsx-core';
import { describe, expect, it } from 'vitest';
import { drawingKeyDown } from './drawing-keys.js';
import { anchorBox, boxAnchor, type DrawingLayer } from './drawings.js';
import { createTestContext } from './test-context.js';
import { closeParens } from './formula-text.js';
import { currentRegionOrAll, fillPlan } from './grid-commands.js';
import { fillTarget } from './grid-pointer.js';

const range = (r1: number, c1: number, r2: number, c2: number) => ({
	start: { row: r1, col: c1 },
	end: { row: r2, col: c2 },
});

describe('grid helpers', () => {
	it('plans Fill Down / Fill Right like Excel', () => {
		expect(fillPlan(range(0, 0, 3, 1), 'down')).toEqual({
			source: range(0, 0, 0, 1),
			target: range(0, 0, 3, 1),
		});
		expect(fillPlan(range(4, 0, 4, 2), 'down')).toEqual({
			source: range(3, 0, 3, 2),
			target: range(3, 0, 4, 2),
		});
		expect(fillPlan(range(0, 0, 0, 0), 'down')).toBeUndefined();
		expect(fillPlan(range(0, 2, 1, 2), 'right')?.source).toEqual(range(0, 1, 1, 1));
	});

	it('extends a fill-handle drag along the dominant axis', () => {
		expect(fillTarget(range(0, 0, 1, 0), { row: 5, col: 1 })).toEqual(range(0, 0, 5, 0));
		expect(fillTarget(range(2, 2, 2, 2), { row: 2, col: 0 })).toEqual(range(2, 0, 2, 2));
		expect(fillTarget(range(0, 0, 1, 1), { row: 1, col: 1 })).toBeUndefined();
	});

	it('selects the current region first, then everything', () => {
		const wb = createWorkbook();
		const sheet = wb.sheets[0]!;
		for (let r = 0; r < 3; r++) for (let c = 0; c < 2; c++) putCell(sheet, r, c, { value: r + c });
		const at = { row: 1, col: 1 };
		const region = currentRegionOrAll(sheet, {
			sheet: 0,
			active: at,
			anchor: at,
			ranges: [range(1, 1, 1, 1)],
		});
		expect(region).toEqual(range(0, 0, 2, 1));
		expect(
			currentRegionOrAll(sheet, { sheet: 0, active: at, anchor: at, ranges: [region!] }),
		).toBeUndefined();
	});

	it('round-trips drawing anchors through plane boxes', () => {
		const wb = createWorkbook();
		const metrics = createGridMetrics(wb.sheets[0]!, { zoom: 100 });
		const anchor = {
			from: { row: 2, col: 1, rowOffset: 9525 * 4, colOffset: 9525 * 10 },
			ext: { cx: 9525 * 200, cy: 9525 * 100 },
		};
		const box = anchorBox(metrics, anchor);
		expect(box.w).toBe(200);
		expect(box.x).toBe(metrics.colLeft(1) + 10);
		expect(boxAnchor(metrics, anchor, box)).toEqual(anchor);
	});

	it('deletes, deselects and nudges the selected drawing from the keyboard', () => {
		const wb = createWorkbook();
		const sheet = wb.sheets[0]!;
		putCell(sheet, 0, 0, { value: 4 });
		const chart: ChartObject = {
			kind: 'chart',
			anchor: { from: { row: 0, col: 2, rowOffset: 0, colOffset: 0 }, ext: { cx: 100, cy: 100 } },
			chartType: 'column',
			series: [],
			showLegend: false,
		};
		sheet.drawings.push(chart);
		const ctx = createTestContext(wb);
		const nudges: number[][] = [];
		const layer = { nudge: (dx: number, dy: number) => nudges.push([dx, dy]) > 0 };
		const key = (k: string, init: KeyboardEventInit = {}) =>
			drawingKeyDown(
				ctx,
				layer as unknown as DrawingLayer,
				new KeyboardEvent('keydown', { key: k, ...init }),
			);
		expect(key('Delete')).toBe(false);
		ctx.selection.set({ drawing: 0 });
		expect(key('ArrowRight')).toBe(true);
		expect(key('ArrowUp', { ctrlKey: true })).toBe(true);
		expect(nudges).toEqual([
			[4, 0],
			[0, -1],
		]);
		expect(key('Escape')).toBe(true);
		expect(ctx.selection.get().drawing).toBeUndefined();
		ctx.selection.set({ drawing: 0 });
		expect(key('Delete')).toBe(true);
		expect(wb.sheets[0]!.drawings).toHaveLength(0);
		expect(ctx.session()?.undoLabel()).toBeTruthy();
		ctx.session()?.undo();
		expect(wb.sheets[0]!.drawings).toHaveLength(1);
	});

	it('closes missing parentheses on commit', () => {
		expect(closeParens('=SUM(A1:A3')).toBe('=SUM(A1:A3)');
		expect(closeParens('=IF(A1,"(",MAX(1')).toBe('=IF(A1,"(",MAX(1))');
		expect(closeParens('text (')).toBe('text (');
	});
});
