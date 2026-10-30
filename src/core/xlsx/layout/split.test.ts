import { describe, expect, it } from 'vitest';
import { createWorksheet } from '../workbook';
import { createGridMetrics } from './metrics';
import {
	clampSplitBar,
	pxToSplit,
	revealInSplit,
	splitAtCell,
	splitToPx,
	splitVisible,
	type SplitLayout,
} from './split';

const metrics = (zoom = 100) => createGridMetrics(createWorksheet('S', 1), { zoom });

describe('split sizes', () => {
	it('converts twentieths of a point to pixels and back', () => {
		expect(splitToPx(3000, 100)).toBe(200);
		expect(splitToPx(3000, 200)).toBe(400);
		expect(pxToSplit(200, 100)).toBe(3000);
		expect(pxToSplit(splitToPx(2400, 150), 150)).toBe(2400);
		expect(splitToPx(0, 100)).toBe(0);
	});

	it('keeps a bar clear of both window edges', () => {
		expect(clampSplitBar(2, 500, 100)).toBe(24);
		expect(clampSplitBar(900, 500, 100)).toBe(476);
		expect(clampSplitBar(100, 500, 100)).toBe(100);
	});
});

describe('splitAtCell', () => {
	const view = { width: 800, height: 400 };

	it('splits above and left of the active cell', () => {
		const m = metrics();
		const split = splitAtCell(m, { row: 3, col: 2 }, { left: 0, top: 0 }, view);
		expect(split?.x).toBe(pxToSplit(m.colLeft(2), 100));
		expect(split?.y).toBe(pxToSplit(m.rowTop(3), 100));
		expect(split?.activePane).toBe('bottomRight');
	});

	it('splits only vertically at column A and only horizontally at row 1', () => {
		const m = metrics();
		const rows = splitAtCell(m, { row: 4, col: 0 }, { left: 0, top: 0 }, view);
		expect(rows?.x).toBe(0);
		expect(rows?.y).toBe(pxToSplit(m.rowTop(4), 100));
		expect(rows?.activePane).toBe('bottomLeft');
		const cols = splitAtCell(m, { row: 0, col: 3 }, { left: 0, top: 0 }, view);
		expect(cols?.y).toBe(0);
		expect(cols?.activePane).toBe('topRight');
	});

	it('measures from the scrolled window, not from the sheet origin', () => {
		const m = metrics();
		const scroll = { left: m.colLeft(5), top: m.rowTop(8) };
		const split = splitAtCell(m, { row: 10, col: 5 }, scroll, view);
		// Nothing left of the cell (it is the first visible column), two rows above it.
		expect(split?.x).toBe(0);
		expect(split?.y).toBe(pxToSplit(m.rowTop(10) - m.rowTop(8), 100));
		expect(split?.topLeft).toEqual({ row: 8, col: 5 });
	});

	it('puts the cross in the middle for A1', () => {
		const split = splitAtCell(metrics(), { row: 0, col: 0 }, { left: 0, top: 0 }, view);
		expect(split?.x).toBe(pxToSplit(400, 100));
		expect(split?.y).toBe(pxToSplit(200, 100));
	});
});

describe('splitVisible and revealInSplit', () => {
	const m = metrics();
	const layout = (over: Partial<SplitLayout> = {}): SplitLayout => ({
		x: m.colLeft(2),
		y: m.rowTop(3),
		left: 0,
		top: 0,
		mainLeft: 0,
		mainTop: 0,
		width: 600,
		height: 300,
		...over,
	});

	it('lets each pane show its own rows and columns', () => {
		const vis = splitVisible(
			m,
			layout({ top: m.rowTop(10), left: m.colLeft(4), mainTop: m.rowTop(50) }),
		);
		expect(vis.frozenRows.slice(0, 3)).toEqual([10, 11, 12]);
		expect(vis.frozenCols.slice(0, 2)).toEqual([4, 5]);
		expect(vis.rows[0]).toBe(50);
		expect(vis.cols[0]).toBe(0);
	});

	it('shows the same cells in both panes right after a split', () => {
		const vis = splitVisible(m, layout());
		expect(vis.frozenRows[0]).toBe(0);
		expect(vis.rows[0]).toBe(0);
		expect(vis.frozenCols[0]).toBe(0);
		expect(vis.cols[0]).toBe(0);
	});

	it('leaves the scroll alone when a pane shows the cell', () => {
		expect(revealInSplit(m, layout(), 1, 1)).toEqual({ left: 0, top: 0 });
	});

	it('scrolls the right and bottom panes to a cell no pane shows', () => {
		const next = revealInSplit(m, layout(), 200, 30);
		expect(next.top).toBeGreaterThan(0);
		expect(next.left).toBeGreaterThan(0);
	});
});
