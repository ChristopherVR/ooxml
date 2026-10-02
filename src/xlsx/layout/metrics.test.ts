import { describe, expect, it } from 'vitest';
import { MAX_ROW } from '../address.js';
import { putCell } from '../cells.js';
import { createWorksheet } from '../workbook.js';
import { AxisMetrics } from './axis-metrics.js';
import { createGridMetrics, zoomPercent } from './metrics.js';
import { visibleCells } from './viewport.js';

/** Brute-force reference for an axis. */
function reference(defaultSize: number, count: number, overrides: Map<number, number>) {
	const sizes = Array.from({ length: count }, (_, i) => overrides.get(i) ?? defaultSize);
	const starts = [0];
	for (const s of sizes) starts.push((starts.at(-1) ?? 0) + s);
	return { sizes, starts };
}

describe('AxisMetrics', () => {
	const overrides = new Map<number, number>([
		[0, 30],
		[3, 0],
		[4, 0],
		[7, 50],
		[12, 5],
		[19, 0],
	]);
	const axis = new AxisMetrics(20, 20, overrides);
	const ref = reference(20, 20, overrides);

	it('matches a brute-force prefix sum', () => {
		for (let i = 0; i <= 20; i++) expect(axis.start(i)).toBe(ref.starts[i]);
		for (let i = 0; i < 20; i++) expect(axis.size(i)).toBe(ref.sizes[i]);
	});

	it('finds the line under every offset, skipping hidden lines', () => {
		const total = ref.starts[20] ?? 0;
		for (let y = 0; y < total; y++) {
			const expected = ref.sizes.findIndex(
				(s, i) => s > 0 && y >= (ref.starts[i] ?? 0) && y < (ref.starts[i] ?? 0) + s,
			);
			expect(axis.at(y)).toBe(expected);
		}
	});

	it('clamps offsets outside the axis', () => {
		expect(axis.at(-10)).toBe(0);
		expect(axis.at(1e9)).toBe(18);
		expect(axis.total).toBe(ref.starts[20]);
	});

	it('reports visibility helpers', () => {
		expect(axis.firstVisible(3)).toBe(5);
		expect(axis.lastVisible(4)).toBe(2);
		expect(axis.lastCustom).toBe(19);
	});

	it('handles an axis with no overrides', () => {
		const plain = new AxisMetrics(10, 100, []);
		expect(plain.start(50)).toBe(500);
		expect(plain.at(505)).toBe(50);
		expect(plain.total).toBe(1000);
	});

	it('stays fast and small on a million lines', () => {
		const sparse = new AxisMetrics(20, MAX_ROW + 1, [
			[500_000, 0],
			[999_999, 100],
		]);
		expect(sparse.start(500_001)).toBe(500_001 * 20 - 20);
		expect(sparse.at(500_000 * 20)).toBe(500_001);
		expect(sparse.at(sparse.start(999_999) + 99)).toBe(999_999);
		expect(sparse.total).toBe((MAX_ROW + 1) * 20 - 20 + 80);
	});
});

describe('createGridMetrics', () => {
	it('uses Excel defaults: 64 px columns and 20 px rows', () => {
		const m = createGridMetrics(createWorksheet('S', 1));
		expect(m.colWidth(0)).toBe(64);
		expect(m.rowHeight(0)).toBe(20);
		expect(m.colLeft(3)).toBe(192);
		expect(m.rowTop(10)).toBe(200);
		expect(m.zoom).toBe(100);
	});

	it('applies custom widths, hidden columns and row heights', () => {
		const sheet = createWorksheet('S', 1);
		sheet.columns = [
			{ min: 1, max: 2, width: 20.7109375, customWidth: true },
			{ min: 4, max: 4, hidden: true },
		];
		sheet.rowInfo.set(2, { height: 30, customHeight: true });
		sheet.rowInfo.set(5, { hidden: true });
		const m = createGridMetrics(sheet);
		expect(m.colWidth(1)).toBe(145);
		expect(m.colWidth(2)).toBe(145);
		expect(m.colWidth(4)).toBe(0);
		expect(m.isColHidden(4)).toBe(true);
		expect(m.colLeft(5)).toBe(64 + 145 + 145 + 64);
		expect(m.colAt(64 + 145 + 145 + 64)).toBe(5);
		expect(m.rowHeight(2)).toBe(40);
		expect(m.rowHeight(5)).toBe(0);
		expect(m.isRowHidden(5)).toBe(true);
		expect(m.rowTop(6)).toBe(20 * 4 + 40);
		expect(m.rowAt(20 * 4 + 40)).toBe(6);
	});

	it('scales by zoom (percent or factor)', () => {
		const sheet = createWorksheet('S', 1);
		expect(createGridMetrics(sheet, { zoom: 200 }).colWidth(0)).toBe(128);
		expect(createGridMetrics(sheet, { zoom: 1.5 }).rowHeight(0)).toBe(30);
		sheet.view.zoom = 50;
		const half = createGridMetrics(sheet);
		expect(half.zoom).toBe(50);
		expect(half.colWidth(0)).toBe(32);
		expect(zoomPercent(1000)).toBe(400);
		expect(zoomPercent(undefined)).toBe(100);
	});

	it('uses the sheet default column width and row height', () => {
		const sheet = createWorksheet('S', 1);
		sheet.defaultColWidth = 13;
		sheet.defaultRowHeight = 18;
		const m = createGridMetrics(sheet);
		expect(m.colWidth(10)).toBe(91);
		expect(m.rowHeight(10)).toBe(24);
	});

	it('computes the scroll extent from the used range plus margins', () => {
		const sheet = createWorksheet('S', 1);
		putCell(sheet, 99, 9, { value: 1 });
		const m = createGridMetrics(sheet, { extraRows: 10, extraCols: 5 });
		expect(m.lastRow).toBe(109);
		expect(m.lastCol).toBe(14);
		expect(m.totalHeight()).toBe(110 * 20);
		expect(m.totalWidth()).toBe(15 * 64);
		expect(m.totalWidth(0)).toBe(64);
	});
});

describe('visibleCells', () => {
	const sheet = createWorksheet('S', 1);
	sheet.rowInfo.set(3, { hidden: true });
	const m = createGridMetrics(sheet);

	it('lists rows and columns in view', () => {
		const v = visibleCells(m, { scrollLeft: 0, scrollTop: 0, width: 200, height: 100 });
		expect(v.rows).toEqual([0, 1, 2, 4, 5]);
		expect(v.cols).toEqual([0, 1, 2, 3]);
		expect(v.frozenRows).toEqual([]);
	});

	it('starts at the scrolled offset', () => {
		const v = visibleCells(m, { scrollLeft: 100, scrollTop: 30, width: 64, height: 40 });
		expect(v.rows).toEqual([1, 2, 4]);
		expect(v.cols).toEqual([1, 2]);
	});

	it('pins frozen panes and scrolls only the rest', () => {
		const v = visibleCells(
			m,
			{ scrollLeft: 0, scrollTop: 20, width: 192, height: 100 },
			{ rows: 2, cols: 1 },
		);
		expect(v.frozenRows).toEqual([0, 1]);
		expect(v.frozenCols).toEqual([0]);
		expect(v.rows).toEqual([4, 5, 6]);
		expect(v.cols).toEqual([1, 2]);
	});

	it('returns nothing for an empty viewport', () => {
		const v = visibleCells(m, { scrollLeft: 0, scrollTop: 0, width: 0, height: 0 });
		expect(v.rows).toEqual([]);
		expect(v.cols).toEqual([]);
	});

	it('works deep in a million-row sheet', () => {
		const v = visibleCells(m, { scrollLeft: 0, scrollTop: 20 * 1_000_000, width: 64, height: 40 });
		expect(v.rows).toEqual([1_000_001, 1_000_002]);
	});
});
