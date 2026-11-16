import { describe, expect, it } from 'vitest';

import {
	selectionScrollTarget,
	selectionZoomBounds,
	zoomToSelectionScale,
} from './zoom-to-selection';
import { MAX_ZOOM_SCALE, MIN_ZOOM_SCALE } from './zoom-step';

describe('zoom to selection', () => {
	it('boxes the selection, giving a zero-size object the minimum extent', () => {
		expect(selectionZoomBounds([])).toBeNull();
		expect(
			selectionZoomBounds([
				{ x: 10, y: 20, width: 100, height: 50 },
				{ x: 300, y: 5, width: 0, height: 0 },
			]),
		).toStrictEqual({ minX: 10, minY: 5, maxX: 312, maxY: 70 });
	});

	it('fits the box in the padded viewport, relative to the fit scale, within the zoom range', () => {
		const bounds = { minX: 0, minY: 0, maxX: 200, maxY: 100 };
		expect(zoomToSelectionScale(bounds, { width: 896, height: 896 }, 1)).toBe(4);
		expect(zoomToSelectionScale(bounds, { width: 896, height: 896 }, 2)).toBe(2);
		const tiny = { minX: 0, minY: 0, maxX: 1, maxY: 1 };
		expect(zoomToSelectionScale(tiny, { width: 4000, height: 4000 }, 1)).toBe(MAX_ZOOM_SCALE);
		const huge = { minX: 0, minY: 0, maxX: 1e6, maxY: 1e6 };
		expect(zoomToSelectionScale(huge, { width: 400, height: 400 }, 1)).toBe(MIN_ZOOM_SCALE);
		expect(zoomToSelectionScale(bounds, { width: 896, height: 896 }, 0)).toBe(MAX_ZOOM_SCALE);
	});

	it('centres the box in the wrapper, never scrolling before the start', () => {
		const bounds = { minX: 100, minY: 100, maxX: 300, maxY: 200 };
		expect(
			selectionScrollTarget(bounds, 2, { left: 20, top: 10 }, { width: 400, height: 200 }),
		).toStrictEqual({ left: 220, top: 210 });
		expect(
			selectionScrollTarget(bounds, 0.1, { left: 0, top: 0 }, { width: 400, height: 200 }),
		).toStrictEqual({ left: 0, top: 0 });
	});
});
