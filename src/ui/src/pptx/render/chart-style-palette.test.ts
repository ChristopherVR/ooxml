import type { ChartPptxElement, PptxChartData } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import {
	DEFAULT_CHART_PALETTE,
	getChartStylePalette,
	resolveChartPalette,
	shade,
	withResolvedPalette,
} from './chart-style-palette';
import { buildChartViewModel } from './chart-view-model';
import type { SvgRect } from './chart-view-model';

/** A theme whose accents match no built-in palette. */
const THEME = ['#2E86AB', '#A23B72', '#F18F01', '#C73E1D', '#3B1F2B', '#6A994E'];

function chartData(overrides: Partial<PptxChartData> = {}): PptxChartData {
	return {
		chartType: 'bar',
		categories: ['Q1'],
		series: Array.from({ length: 7 }, (_, i) => ({ name: `S${i + 1}`, values: [10 + i] })),
		...overrides,
	} as PptxChartData;
}

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

describe('resolveChartPalette', () => {
	it('prefers the parsed colour-style palette', () => {
		const data = chartData({ colorPalette: ['#111111'], themeAccentColors: THEME });
		expect(resolveChartPalette(data)).toStrictEqual(['#111111']);
	});

	it('falls back to the chart style over the theme accents', () => {
		const data = chartData({ style: { styleId: 2 }, themeAccentColors: THEME });
		expect(resolveChartPalette(data)).toStrictEqual([...getChartStylePalette(2, THEME)]);
	});

	it('leaves a chart that already has a palette untouched', () => {
		const data = chartData({ colorPalette: ['#111111'] });
		expect(withResolvedPalette(data)).toBe(data);
	});
});

describe('buildChartViewModel palette', () => {
	it('paints a style-2 chart with no colour part in the theme accents, for every binding', () => {
		const element = {
			id: 'c',
			type: 'chart',
			x: 0,
			y: 0,
			width: 600,
			height: 300,
			chartData: chartData({ style: { styleId: 2 }, themeAccentColors: THEME }),
		} as ChartPptxElement;
		const fills = buildChartViewModel(element)
			.primitives.filter((p): p is SvgRect => p.kind === 'rect' && p.part?.role === 'dataPoint')
			.sort((a, b) => (a.part?.seriesIndex ?? 0) - (b.part?.seriesIndex ?? 0))
			.map((rect) => rect.fill);
		expect(fills[0]).toBe(THEME[0]);
		expect(fills[6]).toBe(shade(THEME[0]!, 0.4));
	});
});
