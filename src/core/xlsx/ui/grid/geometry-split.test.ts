import { createGridMetrics, createWorksheet } from '../../index';
import { describe, expect, it } from 'vitest';
import { GridGeometry, scrollToReveal } from './geometry';

const m = createGridMetrics(createWorksheet('S', 1), { zoom: 100 });

const split = (over: Partial<ConstructorParameters<typeof GridGeometry>[0]> = {}): GridGeometry =>
	new GridGeometry({
		metrics: m,
		freeze: undefined,
		headerW: 30,
		headerH: 20,
		width: 530,
		height: 320,
		scrollLeft: 0,
		scrollTop: 0,
		split: { x: 128, y: 40, left: 0, top: 0 },
		...over,
	});

describe('split window geometry', () => {
	it('identifies the pointer pane even when cells occur in several panes', () => {
		const g = split();
		expect(g.paneAt(31, 21)).toBe('topLeft');
		expect(g.paneAt(160, 21)).toBe('topRight');
		expect(g.paneAt(31, 62)).toBe('bottomLeft');
		expect(g.paneAt(160, 62)).toBe('bottomRight');
		expect(split({ split: { x: 0, y: 40, left: 0, top: 0 } }).paneAt(300, 62)).toBe('bottomLeft');
		expect(split({ split: { x: 128, y: 0, left: 0, top: 0 } }).paneAt(160, 90)).toBe('topRight');
		expect(split({ split: undefined, freeze: { rows: 2, cols: 2 } }).paneAt(31, 21)).toBe(
			'topLeft',
		);
		expect(split({ split: undefined }).paneAt(31, 21)).toBeUndefined();
	});
	const scrolled = {
		scrollLeft: 640,
		scrollTop: 400,
		split: { x: 128, y: 40, left: 64, top: 20 },
	};

	it('lays out four panes that scroll independently', () => {
		const g = split(scrolled);
		expect(g.box('corner')).toEqual({ x: 30, y: 20, w: 128, h: 40 });
		expect(g.box('main')).toEqual({ x: 158, y: 60, w: 372, h: 260 });
		expect(g.origin('corner')).toEqual({ ox: 64, oy: 20 });
		expect(g.origin('top')).toEqual({ ox: 640, oy: 20 });
		expect(g.origin('left')).toEqual({ ox: 64, oy: 400 });
		expect(g.origin('main')).toEqual({ ox: 640, oy: 400 });
	});

	it('hit tests each pane against its own scroll', () => {
		const g = split(scrolled);
		expect(g.hit(31, 21)).toEqual({ area: 'cell', row: m.rowAt(20), col: m.colAt(64) });
		expect(g.hit(160, 62)).toEqual({ area: 'cell', row: m.rowAt(400), col: m.colAt(640) });
	});

	it('does not treat the split as frozen rows and columns', () => {
		const g = split();
		expect(g.fRows).toBe(0);
		expect(g.fCols).toBe(0);
	});

	it('reveals through the panes, scrolling only a cell nobody shows', () => {
		expect(scrollToReveal(split(), 1, 1)).toEqual({ left: 0, top: 0 });
		expect(scrollToReveal(split(), 300, 40).top).toBeGreaterThan(0);
	});
});
