import { describe, expect, it } from 'vitest';
import { isCornerHandle, lockResizeAspect } from './resize-aspect';
import type { BoxTransform, InteractionBox } from './element-transform';
const start: InteractionBox = { x: 0, y: 0, width: 100, height: 50, rotation: 0 };

describe('isCornerHandle', () => {
	it('is true only for the four corners', () => {
		expect((['nw', 'ne', 'se', 'sw'] as const).every(isCornerHandle)).toBeTruthy();
		expect((['n', 'e', 's', 'w'] as const).some(isCornerHandle)).toBeFalsy();
	});
});

describe('lockResizeAspect', () => {
	it('locks a corner resize to the start aspect (larger axis wins)', () => {
		// se: width doubled (scaleW=2), height barely grown (scaleH=1.2) -> scale 2.
		const resized: BoxTransform = { x: 0, y: 0, width: 200, height: 60, rotation: 0 };
		const out = lockResizeAspect(resized, start, 'se');
		expect(out.width).toBe(200);
		expect(out.height).toBe(100);
		expect(out.x).toBe(0);
		expect(out.y).toBe(0);
	});

	it('keeps the anchored (opposite) corner fixed for a nw handle', () => {
		// nw handle anchors the se corner at (100, 50).
		const resized: BoxTransform = { x: -100, y: -25, width: 200, height: 75, rotation: 0 };
		const out = lockResizeAspect(resized, start, 'nw');
		// se corner stays put: x + width === 100, y + height === 50.
		expect(out.x + out.width).toBeCloseTo(100);
		expect(out.y + out.height).toBeCloseTo(50);
	});

	it('returns edge-handle resizes unchanged', () => {
		const resized: BoxTransform = { x: 0, y: 0, width: 300, height: 50, rotation: 0 };
		expect(lockResizeAspect(resized, start, 'e')).toBe(resized);
	});

	it('clamps to the minimum size while preserving the ratio', () => {
		// Shrink toward zero: the min-size clamp bumps both axes proportionally.
		const resized: BoxTransform = { x: 0, y: 0, width: 2, height: 1, rotation: 0 };
		const out = lockResizeAspect(resized, start, 'se', 20);
		expect(out.width).toBeGreaterThanOrEqual(20);
		expect(out.height).toBeGreaterThanOrEqual(10);
		// Ratio preserved (100:50 == 2:1).
		expect(out.width / out.height).toBeCloseTo(2);
	});
});
