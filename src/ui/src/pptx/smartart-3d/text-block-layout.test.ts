import { describe, expect, it } from 'vitest';

import type { SmartArt3DTextBlock } from '../render/smartart-3d-types';
import { layoutTextBlock, MAX_TEXTURE_SIDE, textureScale } from './text-block-layout';

const block = (over: Partial<SmartArt3DTextBlock> = {}): SmartArt3DTextBlock => ({
	lines: [{ text: 'Hello', dy: 0 }],
	x: 0,
	y: 0,
	z: 0,
	maxWidth: 100,
	maxHeight: 40,
	color: '#000',
	fontSize: 16,
	...over,
});

describe('textureScale', () => {
	it('is at least 3x and at most 6x, following the device pixel ratio', () => {
		expect(textureScale(1)).toBe(3);
		expect(textureScale(2)).toBe(4);
		expect(textureScale(3)).toBe(6);
		expect(textureScale(8)).toBe(6);
		expect(textureScale(Number.NaN)).toBe(3);
	});
});

describe('layoutTextBlock', () => {
	it('keeps the text box size when the text fits', () => {
		const layout = layoutTextBlock(block(), () => 40, 1);
		expect(layout.worldWidth).toBe(100);
		expect(layout.worldHeight).toBe(40);
	});

	it('grows symmetrically past a text box that is narrower than its text', () => {
		const layout = layoutTextBlock(block({ maxWidth: 30 }), () => 90, 1);
		// half the measured width plus a quarter font size of padding, each side
		expect(layout.worldWidth).toBe(90 + 2 * 16 * 0.25);
		expect(layout.pixelWidth).toBe(Math.round(layout.worldWidth * 3));
	});

	it('grows to hold lines spread past the box height', () => {
		const lines = [
			{ text: 'a', dy: 40 },
			{ text: 'b', dy: -40 },
		];
		const layout = layoutTextBlock(block({ lines, maxHeight: 20 }), () => 10, 1);
		expect(layout.worldHeight).toBe(2 * (40 + 16 * (0.75 + 0.25)));
	});

	it('ignores empty lines', () => {
		const lines = [{ text: '', dy: 500 }];
		expect(layoutTextBlock(block({ lines }), () => 0, 1).worldHeight).toBe(40);
	});

	it('lowers the scale rather than exceeding the texture cap', () => {
		const layout = layoutTextBlock(block({ maxWidth: 3000 }), () => 10, 3);
		expect(layout.pixelWidth).toBeLessThanOrEqual(MAX_TEXTURE_SIDE);
		expect(layout.scale).toBeLessThan(6);
	});
});
