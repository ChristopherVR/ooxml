/**
 * A chart with no colour-style part paints its chart style over the deck
 * theme's accents (`themeAccentColors`), resolved once by the shared
 * `resolveChartPalette`. Each binding used to resolve its own palette from the
 * Office theme's fixed accents, so a themed deck's charts painted the wrong
 * colours (caught by `e2e/pptx/chart-style-palette.spec.ts`).
 */
import type { PptxChartData } from 'pptx-viewer-core';
import { describe, expect, it } from 'vitest';

import { resolveChartPalette } from './chart-view';

const THEME = ['#2E86AB', '#A23B72', '#F18F01', '#C73E1D', '#3B1F2B', '#6A994E'];

const chartData = {
	chartType: 'bar',
	categories: ['Q1'],
	series: Array.from({ length: 7 }, (_, i) => ({ name: `S${i + 1}`, values: [10 + i] })),
	style: { styleId: 2 },
	themeAccentColors: THEME,
} as PptxChartData;

describe('resolveChartPalette (svelte): theme accents', () => {
	it('paints a style-2 chart in the theme accents, then a darker accent1', () => {
		const palette = resolveChartPalette(chartData);
		expect(palette.slice(0, 6)).toStrictEqual(THEME);
		expect(palette[6]).not.toBe(THEME[0]);
	});
});
