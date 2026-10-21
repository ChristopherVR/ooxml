import { describe, expect, it } from 'vitest';

import { quickStyleSwatchCss, SHAPE_QUICK_STYLES } from './shape-quick-styles';

const HEX = /^#[0-9A-F]{6}$/iu;

describe('shape quick style presets', () => {
	it('store every colour as hex, with transparency in its own field', () => {
		// The writer puts a colour into `a:srgbClr/@val` as it is, so `rgba(...)` there makes an
		// invalid package.
		for (const { name, style } of SHAPE_QUICK_STYLES) {
			for (const key of ['fillColor', 'strokeColor', 'shadowColor'] as const) {
				const color = style[key];
				if (color !== undefined && color !== 'transparent') {
					expect(color, `${name} ${key}`).toMatch(HEX);
				}
			}
			if (style.shadowColor) {
				expect(style.shadowOpacity, `${name} shadowOpacity`).toBeGreaterThan(0);
			}
		}
	});

	it('preview a shadow at its own opacity', () => {
		const blueShadow = SHAPE_QUICK_STYLES.find((preset) => preset.name === 'Blue Shadow')!;
		expect(quickStyleSwatchCss(blueShadow).boxShadow).toBe('2px 2px 6px rgba(0, 0, 0, 0.35)');
	});
});
