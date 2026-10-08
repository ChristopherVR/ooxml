import { describe, expect, it } from 'vitest';

import { DEFAULT_CHART_PALETTE, getChartStylePalette } from './chart-helpers';
import { shade, tint } from './chart-palette';

const ACCENTS = ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47'];
const ramp = (base: string) => [
	shade(base, 0.5),
	shade(base, 0.35),
	shade(base, 0.15),
	base,
	tint(base, 0.2),
	tint(base, 0.4),
	tint(base, 0.6),
	tint(base, 0.8),
];

describe('getChartStylePalette (Office chart styles 1-48)', () => {
	it('maps style 2, the Office default, to the accents in theme order from accent 1', () => {
		const palette = getChartStylePalette(2);
		expect(palette.slice(0, 6)).toStrictEqual(ACCENTS);
		// Series 7-12 reuse the accents, darker.
		expect(palette.slice(6)).toStrictEqual(ACCENTS.map((accent) => shade(accent, 0.4)));
	});

	it('maps style 1 to greyscale and styles 3-8 to monochrome accent 1 to accent 6', () => {
		expect(getChartStylePalette(1)).toStrictEqual(ramp('#7F7F7F'));
		expect(getChartStylePalette(3)).toStrictEqual(ramp(ACCENTS[0]!));
		expect(getChartStylePalette(4)).toStrictEqual(ramp(ACCENTS[1]!));
		expect(getChartStylePalette(8)).toStrictEqual(ramp(ACCENTS[5]!));
	});

	it('repeats the colour columns in every gallery row (rows differ in effects only)', () => {
		expect(getChartStylePalette(10)).toStrictEqual(getChartStylePalette(2));
		expect(getChartStylePalette(10)[0]).toBe('#4472C4');
		expect(getChartStylePalette(48)).toStrictEqual(ramp(ACCENTS[5]!));
		for (let row = 1; row < 6; row++)
			for (let column = 1; column <= 8; column++)
				expect(getChartStylePalette(row * 8 + column)).toStrictEqual(getChartStylePalette(column));
	});

	it('falls back to the default palette for a missing or out-of-range style', () => {
		expect(getChartStylePalette(undefined)).toBe(DEFAULT_CHART_PALETTE);
		expect(getChartStylePalette(0)).toBe(DEFAULT_CHART_PALETTE);
		expect(getChartStylePalette(49)).toBe(DEFAULT_CHART_PALETTE);
	});
});
