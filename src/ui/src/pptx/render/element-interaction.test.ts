import { describe, expect, it } from 'vitest';

import {
	computeMarqueeHitIds,
	computeRotation,
	isAdditiveSelectionPress,
	mergeAdditiveSelection,
	snapBoxToGrid,
} from './element-interaction';
import type { MarqueeElementRect, MarqueeRect } from './element-interaction';

describe('computeRotation', () => {
	const center = { x: 0, y: 0 };

	it('returns 0 when the pointer is directly above the center', () => {
		expect(computeRotation(center, { x: 0, y: -10 })).toBeCloseTo(0, 6);
	});

	it('returns 90 to the right', () => {
		expect(computeRotation(center, { x: 10, y: 0 })).toBeCloseTo(90, 6);
	});

	it('returns 180 below', () => {
		expect(computeRotation(center, { x: 0, y: 10 })).toBeCloseTo(180, 6);
	});

	it('returns 270 to the left', () => {
		expect(computeRotation(center, { x: -10, y: 0 })).toBeCloseTo(270, 6);
	});

	it('normalises to [0, 360)', () => {
		const angle = computeRotation(center, { x: -1, y: -10 });
		expect(angle).toBeGreaterThanOrEqual(0);
		expect(angle).toBeLessThan(360);
	});
});

describe('snapBoxToGrid', () => {
	const start = { x: 100, y: 100, width: 200, height: 150 };

	it('snaps the SE right/bottom edges keeping the origin fixed', () => {
		// Right edge 100+253 -> nearest 10 = 350 -> width 250; bottom 100+197 ->
		// nearest 10 = 300 -> height 200.
		const r = snapBoxToGrid({ ...start, width: 253, height: 197 }, 'se', 10);
		expect(r.x).toBe(100);
		expect(r.y).toBe(100);
		expect(r.width).toBe(250);
		expect(r.height).toBe(200);
	});

	it('snaps the NW origin growing the dimension to compensate', () => {
		// origin 87/83 -> nearest 10 = 90/80.
		const r = snapBoxToGrid({ x: 87, y: 83, width: 213, height: 217 }, 'nw', 10);
		expect(r.x).toBe(90);
		expect(r.y).toBe(80);
	});

	it('returns the box unchanged for a null handle or non-positive grid', () => {
		expect(snapBoxToGrid(start, null, 10)).toBe(start);
		expect(snapBoxToGrid(start, 'se', 0)).toBe(start);
	});

	it('clamps a snapped dimension to the min size', () => {
		const r = snapBoxToGrid({ x: 0, y: 0, width: 1, height: 1 }, 'se', 10, 8);
		expect(r.width).toBeGreaterThanOrEqual(8);
		expect(r.height).toBeGreaterThanOrEqual(8);
	});
});

describe('computeMarqueeHitIds', () => {
	const elements: MarqueeElementRect[] = [
		{ id: 'a', x: 10, y: 10, width: 50, height: 50 },
		{ id: 'b', x: 100, y: 100, width: 50, height: 50 },
		{ id: 'c', x: 200, y: 200, width: 50, height: 50 },
	];

	it('returns empty for a tiny marquee (both dims <= 3)', () => {
		const m: MarqueeRect = { startX: 10, startY: 10, currentX: 12, currentY: 12 };
		expect(computeMarqueeHitIds(m, elements)).toStrictEqual([]);
	});

	it('selects elements inside the rectangle', () => {
		const m: MarqueeRect = { startX: 0, startY: 0, currentX: 70, currentY: 70 };
		expect(computeMarqueeHitIds(m, elements)).toStrictEqual(['a']);
	});

	it('normalises reversed corners', () => {
		const m: MarqueeRect = { startX: 70, startY: 70, currentX: 0, currentY: 0 };
		expect(computeMarqueeHitIds(m, elements)).toStrictEqual(['a']);
	});

	it('clamps small elements to the min size so they stay selectable', () => {
		const tiny: MarqueeElementRect = { id: 'tiny', x: 50, y: 50, width: 2, height: 2 };
		const m: MarqueeRect = { startX: 45, startY: 45, currentX: 63, currentY: 63 };
		expect(computeMarqueeHitIds(m, [tiny], 12)).toStrictEqual(['tiny']);
	});
});

describe('mergeAdditiveSelection', () => {
	it('merges and de-duplicates, base ids first', () => {
		expect(mergeAdditiveSelection(['a', 'b'], ['b', 'c'])).toStrictEqual(['a', 'b', 'c']);
	});

	it('handles an undefined base', () => {
		expect(mergeAdditiveSelection(undefined, ['a', 'b'])).toStrictEqual(['a', 'b']);
	});
});

describe('isAdditiveSelectionPress', () => {
	it('treats a plain press as a replacement', () => {
		expect(
			isAdditiveSelectionPress({ shiftKey: false, ctrlKey: false, metaKey: false }),
		).toBeFalsy();
	});

	/**
	 * PowerPoint extends a selection with Shift OR Ctrl (Cmd on macOS). Vanilla
	 * consulted `shiftKey` alone, so Ctrl+click on Windows quietly threw away a
	 * multi-selection the user was part-way through building.
	 */
	it('treats Shift, Ctrl and Meta alike', () => {
		expect(isAdditiveSelectionPress({ shiftKey: true })).toBeTruthy();
		expect(isAdditiveSelectionPress({ shiftKey: false, ctrlKey: true })).toBeTruthy();
		expect(isAdditiveSelectionPress({ shiftKey: false, metaKey: true })).toBeTruthy();
	});
});
