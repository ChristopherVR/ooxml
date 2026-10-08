import { describe, it, expect } from 'vitest';

import { getChartStylePalette, DEFAULT_CHART_PALETTE, tint, shade } from './chart-style-palettes';

describe('getChartStylePalette', () => {
	it('should return the default palette when styleId is undefined', () => {
		expect(getChartStylePalette(undefined)).toBe(DEFAULT_CHART_PALETTE);
	});

	it('should return the default palette when styleId is 0', () => {
		expect(getChartStylePalette(0)).toBe(DEFAULT_CHART_PALETTE);
	});

	it('should return the default palette when styleId is negative', () => {
		expect(getChartStylePalette(-1)).toBe(DEFAULT_CHART_PALETTE);
	});

	it('should return the default palette when styleId is > 48', () => {
		expect(getChartStylePalette(49)).toBe(DEFAULT_CHART_PALETTE);
	});

	it('should return an 8-colour array for style 1', () => {
		const palette = getChartStylePalette(1);
		expect(palette).toHaveLength(8);
		palette.forEach((c) => {
			expect(c).toMatch(/^#[0-9a-f]{6}$/iu);
		});
	});

	it('should return different palettes for different style IDs', () => {
		const p1 = getChartStylePalette(1);
		const p2 = getChartStylePalette(2);
		const p3 = getChartStylePalette(3);
		// Office's gallery columns: 1 greyscale, 2 colourful, 3-8 monochrome accent 1-6
		expect(p1).not.toStrictEqual(p2);
		expect(p2).not.toStrictEqual(p3);
	});

	it('should return the same palette for the same style ID (cached)', () => {
		const a = getChartStylePalette(5);
		const b = getChartStylePalette(5);
		expect(a).toBe(b); // exact same reference (cached)
	});

	it('should return valid hex colours for all style IDs 1-48', () => {
		for (let id = 1; id <= 48; id++) {
			const palette = getChartStylePalette(id);
			expect(palette.length).toBeGreaterThanOrEqual(8);
			palette.forEach((c) => {
				expect(c).toMatch(/^#[0-9a-f]{6}$/iu);
			});
		}
	});

	it('style 1 should be the greyscale ramp', () => {
		expect(getChartStylePalette(1)[3]).toBe('#7F7F7F');
	});

	it('style 2 (the Office default) should start with accent1 (Office blue)', () => {
		const palette = getChartStylePalette(2);
		expect(palette.slice(0, 6)).toStrictEqual([
			'#4472C4',
			'#ED7D31',
			'#A5A5A5',
			'#FFC000',
			'#5B9BD5',
			'#70AD47',
		]);
	});

	it('style 3 should be a monochromatic ramp from accent1, style 4 from accent2', () => {
		// Monochromatic ramp: darkest to lightest, the accent in the middle
		expect(getChartStylePalette(3)[3]).toBe('#4472C4');
		expect(getChartStylePalette(4)[3]).toBe('#ED7D31');
		expect(getChartStylePalette(3)).toHaveLength(8);
	});

	it('every gallery row repeats the colour columns (rows differ in effects only)', () => {
		expect(getChartStylePalette(9)).toStrictEqual(getChartStylePalette(1));
		expect(getChartStylePalette(10)).toStrictEqual(getChartStylePalette(2));
		expect(getChartStylePalette(48)).toStrictEqual(getChartStylePalette(8));
	});
});

describe('tint', () => {
	it('should return white when amount is 1', () => {
		expect(tint('#000000', 1)).toBe('#ffffff');
	});

	it('should return the original colour when amount is 0', () => {
		expect(tint('#4472C4', 0)).toBe('#4472c4');
	});

	it('should lighten a colour', () => {
		const result = tint('#4472C4', 0.5);
		// Should be lighter than original, check the R channel
		const r = parseInt(result.slice(1, 3), 16);
		expect(r).toBeGreaterThan(0x44);
	});
});

describe('shade', () => {
	it('should return black when amount is 1', () => {
		expect(shade('#ffffff', 1)).toBe('#000000');
	});

	it('should return the original colour when amount is 0', () => {
		expect(shade('#4472C4', 0)).toBe('#4472c4');
	});

	it('should darken a colour', () => {
		const result = shade('#4472C4', 0.5);
		// Should be darker, check the R channel
		const r = parseInt(result.slice(1, 3), 16);
		expect(r).toBeLessThan(0x44);
	});
});
