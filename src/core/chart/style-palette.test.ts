import { describe, expect, it } from 'vitest';

import { DEFAULT_CHART_PALETTE, getChartStylePalette, shade } from './style-palette';

/** A theme whose accents match no built-in palette. */
const THEME = ['#2E86AB', '#A23B72', '#F18F01', '#C73E1D', '#3B1F2B', '#6A994E'];

describe('getChartStylePalette over the deck theme', () => {
	it('builds the colourful style over the theme accents, then darker accents', () => {
		const palette = getChartStylePalette(2, THEME);
		expect(palette.slice(0, 6)).toStrictEqual(THEME);
		expect(palette[6]).toBe(shade(THEME[0]!, 0.4));
	});

	it('builds the monochrome styles from the matching theme accent', () => {
		expect(getChartStylePalette(3, THEME)[3]).toBe(THEME[0]);
		expect(getChartStylePalette(8, THEME)[3]).toBe(THEME[5]);
	});

	it('keeps the two default extras after the theme accents when there is no style', () => {
		expect(getChartStylePalette(undefined, THEME)).toStrictEqual([
			...THEME,
			...DEFAULT_CHART_PALETTE.slice(6),
		]);
	});

	it('ignores accents that are not six #RRGGBB colours', () => {
		expect(getChartStylePalette(2, THEME.slice(0, 5))).toStrictEqual(getChartStylePalette(2));
		expect(getChartStylePalette(undefined, ['red', ...THEME.slice(1)])).toBe(DEFAULT_CHART_PALETTE);
	});
});
