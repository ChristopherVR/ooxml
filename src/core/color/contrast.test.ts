import { describe, expect, it } from 'vitest';
import { contrastRatio, parseHexColor, readableTextColor, relativeLuminance } from './contrast';

describe('contrastRatio', () => {
	it('follows the WCAG definition', () => {
		expect(contrastRatio('#000000', '#ffffff')).toBeCloseTo(21, 5);
		expect(contrastRatio('#fff', '#ffffff')).toBe(1);
		// Published reference values: #767676 is the lightest grey that passes 4.5:1 on white.
		expect(contrastRatio('#767676', '#ffffff')).toBeCloseTo(4.54, 2);
		expect(contrastRatio('#0f6cbd', '#ffffff')).toBeCloseTo(5.36, 1);
	});

	it('is symmetric and rejects colours it cannot read', () => {
		expect(contrastRatio('#ffffff', '#0f6cbd')).toBe(contrastRatio('#0f6cbd', '#ffffff'));
		expect(contrastRatio('red', '#ffffff')).toBeUndefined();
		expect(contrastRatio('#12345', '#ffffff')).toBeUndefined();
	});
});

describe('readableTextColor', () => {
	it('picks the candidate that reads best, at least 4.5:1 for white and black', () => {
		expect(readableTextColor('#ffffff')).toBe('#000000');
		expect(readableTextColor('#0f6cbd')).toBe('#ffffff');
		for (let grey = 0; grey <= 255; grey += 1) {
			const hex = `#${grey.toString(16).padStart(2, '0').repeat(3)}`;
			expect(contrastRatio(readableTextColor(hex), hex)).toBeGreaterThanOrEqual(4.5);
		}
	});

	it('honours custom candidates and falls back to the first for unknown fills', () => {
		expect(readableTextColor('#202228', ['#141414', '#f0f0f0'])).toBe('#f0f0f0');
		expect(readableTextColor('hsl(0 0% 50%)', ['#141414', '#f0f0f0'])).toBe('#141414');
	});
});

describe('parseHexColor and relativeLuminance', () => {
	it('parse short and long forms', () => {
		expect(parseHexColor('#abc')).toEqual([170, 187, 204]);
		expect(parseHexColor(' 0F6CBD ')).toEqual([15, 108, 189]);
		expect(parseHexColor('#0f6cbd80')).toBeNull();
		expect(relativeLuminance(255, 255, 255)).toBeCloseTo(1, 10);
		expect(relativeLuminance(0, 0, 0)).toBe(0);
	});
});
