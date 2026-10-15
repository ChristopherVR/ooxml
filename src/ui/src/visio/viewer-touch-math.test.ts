import { expect, it } from 'vitest';
import {
	classifyTouch,
	clampZoom,
	hitProbes,
	hitTolerance,
	pinchDistance,
	pinchMidpoint,
	pinchScrollDelta,
	pointFraction,
	pinchZoom,
} from './viewer-touch-math';

it('computes pinch distance and midpoint', () => {
	expect(pinchDistance({ x: 0, y: 0 }, { x: 3, y: 4 })).toBe(5);
	expect(pinchMidpoint({ x: 0, y: 10 }, { x: 10, y: 20 })).toEqual({ x: 5, y: 15 });
});
it('scales zoom with finger distance and clamps to the limits', () => {
	expect(pinchZoom(1, 100, 200)).toBe(2);
	expect(pinchZoom(1, 100, 50)).toBe(0.5);
	expect(pinchZoom(4, 100, 1000)).toBe(8);
	expect(pinchZoom(0.2, 100, 10)).toBe(0.1);
	expect(pinchZoom(2, 0, 50)).toBe(2);
	expect(clampZoom(Number.NaN)).toBe(1);
});
it('scrolls the pinched paper point back under the fingers', () => {
	const rect = { left: 100, top: 50, width: 400, height: 200 };
	const fraction = pointFraction({ x: 200, y: 100 }, rect);
	expect(fraction).toEqual({ x: 0.25, y: 0.25 });
	const grown = { left: 20, top: 10, width: 800, height: 400 };
	expect(pinchScrollDelta(fraction, grown, { x: 200, y: 100 })).toEqual({ x: 20, y: 10 });
});
it('classifies tap, long-press, drag and pending', () => {
	expect(classifyTouch({ moved: 2, elapsed: 100, released: true })).toBe('tap');
	expect(classifyTouch({ moved: 2, elapsed: 100, released: false })).toBe('pending');
	expect(classifyTouch({ moved: 2, elapsed: 600, released: false })).toBe('long-press');
	expect(classifyTouch({ moved: 20, elapsed: 600, released: false })).toBe('drag');
});
it('widens hit tolerance for touch only', () => {
	expect(hitTolerance('mouse')).toBe(0);
	expect(hitTolerance('pen')).toBe(0);
	expect(hitTolerance('touch')).toBeGreaterThan(8);
	expect(hitProbes(0)).toHaveLength(1);
	expect(hitProbes(14)).toHaveLength(17);
});
