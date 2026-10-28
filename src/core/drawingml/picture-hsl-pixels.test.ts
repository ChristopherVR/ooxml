import { expect, it } from 'vitest';
import { applyPictureHslPixels } from './picture-hsl-pixels';
it('matches independently inspected Office XML HSL controls and retains alpha', () => {
	for (const [saturation, luminance, expected] of [
		[0.5, 1, [128, 0, 0]],
		[1, 0.5, [191, 64, 64]],
		[1, 1, [255, 0, 0]],
	] as const) {
		const pixels = new Uint8ClampedArray([255, 0, 0, 128]);
		applyPictureHslPixels(pixels, {
			kind: 'standard',
			name: 'hsl',
			values: { saturation, luminance },
		});
		expect([...pixels]).toEqual([...expected, 128]);
	}
	const rgb = new Uint8ClampedArray([255, 0, 0, 255]);
	applyPictureHslPixels(rgb, {
		kind: 'standard',
		name: 'hsl',
		values: { hue: 30, saturation: 0.5, luminance: 0.2 },
	});
	expect([...rgb]).toEqual([77, 64, 51, 255]);
});
it('moves hue toward the specified tint without changing saturation, lightness or alpha', () => {
	const pixels = new Uint8ClampedArray([255, 0, 0, 128]);
	applyPictureHslPixels(pixels, {
		kind: 'standard',
		name: 'tint',
		values: { hue: 60, amount: 0.3 },
	});
	expect([...pixels]).toEqual([255, 77, 0, 128]);
});
