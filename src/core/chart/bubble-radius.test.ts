import { describe, expect, it } from 'vitest';

import { computeBubbleRadius, maxBubbleSize, medianBubbleRadius } from './bubble-radius';

describe('computeBubbleRadius', () => {
	it('uses the median radius for a bubble with no size', () => {
		expect(computeBubbleRadius(undefined, 100, 10)).toBe(10);
		expect(computeBubbleRadius(undefined, 100, 10, { bubbleScale: 50 })).toBe(5);
	});

	it('spans 0.5x to 2x the median linearly by default', () => {
		expect(computeBubbleRadius(0, 100, 10)).toBeCloseTo(5);
		expect(computeBubbleRadius(100, 100, 10)).toBeCloseTo(20);
	});

	it('tracks area with the square root when sizeRepresents is area', () => {
		const half = computeBubbleRadius(25, 100, 10, { sizeRepresents: 'area' });
		expect(half).toBeCloseTo(5 + 0.5 * 15);
	});

	it('sizes a negative bubble by its magnitude and survives a zero maximum', () => {
		expect(computeBubbleRadius(-100, 100, 10)).toBeCloseTo(20);
		expect(Number.isFinite(computeBubbleRadius(5, 0, 10))).toBe(true);
	});
});

describe('bubble envelope', () => {
	it('measures against the largest magnitude, at least 1', () => {
		expect(maxBubbleSize([])).toBe(1);
		expect(maxBubbleSize([0.2, -0.5])).toBe(1);
		expect(maxBubbleSize([3, -9])).toBe(9);
	});

	it('takes 4% of the plot short side as the median radius', () => {
		expect(medianBubbleRadius(400, 200)).toBe(8);
	});
});

describe('bubble envelope over many points', () => {
	it('measures more bubbles than a call can pass as arguments', () => {
		const sizes = Array.from({ length: 200_000 }, (_, i) => i % 50);
		sizes[90_000] = -75;
		expect(maxBubbleSize(sizes)).toBe(75);
	});
});
