import { createGridMetrics, createWorksheet } from '@christophervr/xlsx-core';
import { describe, expect, it } from 'vitest';
import { GridGeometry, rowHeaderWidth, scrollToReveal } from './geometry.js';

const metrics = () => createGridMetrics(createWorksheet('S', 1), { zoom: 100 });

describe('grid geometry', () => {
	it('lays out the four quadrants around frozen panes', () => {
		const m = metrics();
		const g = new GridGeometry({
			metrics: m,
			freeze: { rows: 1, cols: 1 },
			headerW: 30,
			headerH: 20,
			width: 500,
			height: 300,
			scrollLeft: 0,
			scrollTop: 0,
		});
		expect(g.box('corner')).toEqual({ x: 30, y: 20, w: m.colWidth(0), h: m.rowHeight(0) });
		expect(g.box('main').x).toBe(30 + m.colWidth(0));
		expect(g.origin('main')).toEqual({ ox: m.colWidth(0), oy: m.rowHeight(0) });
		expect(g.origin('corner')).toEqual({ ox: 0, oy: 0 });
		expect(g.quadrantOf(0, 3)).toBe('top');
		expect(g.quadrantOf(5, 0)).toBe('left');
	});

	it('maps view pixels to cells and headers, honouring scroll', () => {
		const m = metrics();
		const g = new GridGeometry({
			metrics: m,
			freeze: undefined,
			headerW: 30,
			headerH: 20,
			width: 500,
			height: 300,
			scrollLeft: m.colWidth(0),
			scrollTop: m.rowHeight(0) * 2,
		});
		expect(g.hit(31, 21)).toEqual({ area: 'cell', row: 2, col: 1 });
		expect(g.hit(5, 5).area).toBe('corner');
		expect(g.hit(40, 5)).toEqual({ area: 'col-header', row: -1, col: 1 });
		expect(g.hit(5, 40).area).toBe('row-header');
		expect(g.colX(1)).toBe(30);
		const right = g.colX(1) + m.colWidth(1);
		expect(g.colBorderAt(right)).toBe(1);
		expect(g.colBorderAt(right - 20)).toBeUndefined();
	});

	it('computes the scroll that reveals a cell', () => {
		const m = metrics();
		const g = new GridGeometry({
			metrics: m,
			freeze: undefined,
			headerW: 30,
			headerH: 20,
			width: 330,
			height: 220,
			scrollLeft: 0,
			scrollTop: 0,
		});
		const next = scrollToReveal(g, 40, 10);
		expect(next.top).toBeGreaterThan(0);
		expect(next.left).toBeGreaterThan(0);
		expect(scrollToReveal(g, 0, 0)).toEqual({ left: 0, top: 0 });
		expect(rowHeaderWidth(99, 100)).toBeLessThan(rowHeaderWidth(99_999, 100));
	});
});
