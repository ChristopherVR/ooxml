import { describe, expect, it } from 'vitest';
import { EMU_PER_PIXEL } from '../../units/constants.js';
import type { DrawingAnchor } from '../model.js';
import { createWorksheet } from '../workbook.js';
import {
	DEFAULT_ANCHOR_EXTENT_EMU,
	anchorKind,
	anchorToPixelBox,
	pictureAnchorAt,
	pixelBoxToAnchor,
	pixelSizeToExtent,
} from './anchors.js';
import { createGridMetrics } from './metrics.js';

const E = EMU_PER_PIXEL;
const at = (row: number, col: number, rowPx = 0, colPx = 0) => ({
	row,
	col,
	rowOffset: rowPx * E,
	colOffset: colPx * E,
});

describe('anchorToPixelBox', () => {
	const sheet = createWorksheet('S', 1);

	it('spans the markers of a two-cell anchor (64 x 20 px default cells)', () => {
		const anchor: DrawingAnchor = { from: at(1, 1, 5, 10), to: at(4, 3, 2, 6) };
		expect(anchorToPixelBox(sheet, anchor)).toEqual({ x: 74, y: 25, w: 192 + 6 - 74, h: 82 - 25 });
	});

	it('keeps a collapsed two-cell anchor at least 4 px wide and high', () => {
		const anchor: DrawingAnchor = { from: at(2, 2), to: at(2, 2) };
		expect(anchorToPixelBox(sheet, anchor)).toEqual({ x: 128, y: 40, w: 4, h: 4 });
	});

	it('uses the extent of a one-cell anchor, defaulting to 200 px', () => {
		const anchor: DrawingAnchor = { from: at(0, 1), ext: { cx: 100 * E, cy: 50 * E } };
		expect(anchorToPixelBox(sheet, anchor)).toEqual({ x: 64, y: 0, w: 100, h: 50 });
		const bare = anchorToPixelBox(sheet, { from: at(0, 0) });
		expect(bare.w).toBe(DEFAULT_ANCHOR_EXTENT_EMU / E);
		expect(bare.w).toBe(200);
	});

	it('places absolute anchors (A1 plus offsets) by their position', () => {
		const anchor: DrawingAnchor = {
			from: { row: 0, col: 0, colOffset: 300 * E, rowOffset: 90 * E },
			ext: { cx: 40 * E, cy: 30 * E },
		};
		expect(anchorToPixelBox(sheet, anchor)).toEqual({ x: 300, y: 90, w: 40, h: 30 });
	});

	it('scales offsets and extents by the zoom and follows custom sizes', () => {
		const wide = createWorksheet('W', 1);
		wide.columns = [{ min: 0, max: 0, width: 20.7109375, customWidth: true }];
		wide.rowInfo.set(0, { hidden: true });
		const metrics = createGridMetrics(wide, { zoom: 200 });
		const anchor: DrawingAnchor = { from: at(1, 1, 1, 2), ext: { cx: 10 * E, cy: 10 * E } };
		expect(anchorToPixelBox(wide, anchor, metrics)).toEqual({ x: 290 + 4, y: 2, w: 20, h: 20 });
	});

	it('defaults the metrics to the sheet zoom', () => {
		const half = createWorksheet('H', 1);
		half.view.zoom = 50;
		const box = anchorToPixelBox(half, { from: at(0, 2), ext: { cx: 100 * E, cy: 100 * E } });
		expect(box).toEqual({ x: 64, y: 0, w: 50, h: 50 });
	});
});

describe('pixelBoxToAnchor', () => {
	const sheet = createWorksheet('S', 1);
	const metrics = createGridMetrics(sheet);

	it('builds a one-cell anchor by default', () => {
		expect(pixelBoxToAnchor(sheet, { x: 74, y: 25, w: 100, h: 50 })).toEqual({
			from: at(1, 1, 5, 10),
			ext: { cx: 100 * E, cy: 50 * E },
		});
	});

	it('keeps the form of a previous anchor', () => {
		const previous: DrawingAnchor = { from: at(0, 0), to: at(1, 1) };
		expect(pixelBoxToAnchor(sheet, { x: 74, y: 25, w: 124, h: 57 }, metrics, previous)).toEqual({
			from: at(1, 1, 5, 10),
			to: at(4, 3, 2, 6),
		});
		expect(anchorKind(previous)).toBe('twoCell');
		expect(pixelBoxToAnchor(sheet, { x: 0, y: 0, w: 1, h: 1 }, metrics, 'twoCell').to).toEqual(
			at(0, 0, 1, 1),
		);
	});

	it('clamps negative positions to A1', () => {
		const anchor = pixelBoxToAnchor(sheet, { x: -10, y: -5, w: 20, h: 20 });
		expect(anchor.from.col).toBe(0);
		expect(anchor.from.row).toBe(0);
	});

	it('stores EMU at 100% zoom whatever the view zoom', () => {
		const zoomed = createGridMetrics(sheet, { zoom: 200 });
		const anchor = pixelBoxToAnchor(sheet, { x: 148, y: 50, w: 200, h: 100 }, zoomed);
		expect(anchor).toEqual({ from: at(1, 1, 5, 10), ext: { cx: 100 * E, cy: 50 * E } });
	});

	it('round-trips with anchorToPixelBox for both forms and zooms', () => {
		for (const zoom of [100, 75, 150]) {
			const m = createGridMetrics(sheet, { zoom });
			for (const like of ['oneCell', 'twoCell'] as const) {
				const box = { x: 3 * m.colWidth(0) + 7, y: 2 * m.rowHeight(0) + 3, w: 120, h: 80 };
				const back = anchorToPixelBox(sheet, pixelBoxToAnchor(sheet, box, m, like), m);
				expect(back.x).toBeCloseTo(box.x, 3);
				expect(back.y).toBeCloseTo(box.y, 3);
				expect(back.w).toBeCloseTo(box.w, 3);
				expect(back.h).toBeCloseTo(box.h, 3);
			}
		}
	});

	it('skips hidden columns when choosing the marker cell', () => {
		const hidden = createWorksheet('X', 1);
		hidden.columns = [{ min: 1, max: 1, hidden: true }];
		const anchor = pixelBoxToAnchor(hidden, { x: 70, y: 0, w: 10, h: 10 });
		expect(anchor.from.col).toBe(2);
		expect(anchor.from.colOffset).toBe(6 * E);
	});
});

describe('picture extents', () => {
	it('converts pixels to whole EMU', () => {
		expect(pixelSizeToExtent(1.5, 2)).toEqual({ cx: Math.round(1.5 * E), cy: 2 * E });
	});

	it('anchors a picture at a cell, scaled to 640 px wide at most', () => {
		expect(pictureAnchorAt({ row: 3, col: 2 }, 1280, 720)).toEqual({
			from: at(3, 2),
			ext: { cx: 640 * E, cy: 360 * E },
		});
		expect(pictureAnchorAt({ row: 0, col: 0 }, 100, 50).ext).toEqual({ cx: 100 * E, cy: 50 * E });
		expect(pictureAnchorAt({ row: 0, col: 0 }, 400, 400, 200).ext).toEqual({
			cx: 200 * E,
			cy: 200 * E,
		});
	});
});
