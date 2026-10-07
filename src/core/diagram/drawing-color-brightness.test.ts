import { expect, it } from 'vitest';
import { drawingColorBrightness, withDrawingColorBrightness } from './drawing-color-brightness';
import type { DiagramColor } from './types';

it('retains theme identity and unrelated ordered transforms when replacing brightness', () => {
	const color: DiagramColor = {
		kind: 'scheme',
		value: 'accent1',
		transforms: [
			{ name: 'alpha', value: '63000' },
			{ name: 'lumMod', value: '50000' },
			{ name: 'lumOff', value: '50000' },
			{ name: 'satMod', value: '80000' },
		],
	};
	const before = structuredClone(color);
	const next = withDrawingColorBrightness(color, -42);
	expect(color).toEqual(before);
	expect(next).toEqual({
		kind: 'scheme',
		value: 'accent1',
		transforms: [
			{ name: 'alpha', value: '63000' },
			{ name: 'satMod', value: '80000' },
			{ name: 'lumMod', value: '58000' },
		],
	});
	expect(drawingColorBrightness(next)).toBe(-42);
});

it('reports noncanonical luminance as unknown rather than displaying an incorrect brightness', () => {
	for (const transforms of [
		[
			{ name: 'lumOff', value: '50000' },
			{ name: 'lumMod', value: '50000' },
		],
		[{ name: 'lum', value: '50000' }],
		[
			{ name: 'lumMod', value: '50000' },
			{ name: 'lumOff', value: '10000' },
		],
		[{ name: 'lumMod', value: 'bad' }],
		[{ name: 'lumMod', value: '' }],
		[
			{ name: 'lumMod', value: '50000' },
			{ name: 'lumMod', value: '60000' },
		],
	])
		expect(drawingColorBrightness({ kind: 'srgb', value: 'FF0000', transforms })).toBeUndefined();
});
