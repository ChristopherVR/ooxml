// The neutral entry to the chart painter: a parsed `c:chartSpace` and a DrawingML theme in, a
// standalone SVG document out. Word and PowerPoint draw chart parts through this; Excel builds
// the same summary with its anchor and live cell references.
import { themeSlotHex } from '../../drawingml/theme-color';
import type { ThemeColorScheme } from '../../drawingml/theme-model';
import type { ChartSpace } from '../model';
import { renderChartSvg, type ChartSvgOptions } from './chart-svg';
import { chartSummaryView } from './chart-view';
import { chartSummaryFromSpace, type ChartSummaryExtras } from './summary-from-space';
import type { ChartSummaryType, ChartThemePalette } from './summary';
import { DEFAULT_THEME_COLORS, THEME_SLOTS } from './theme-palette';

/** Chart families the SVG painter draws; the others render a labelled frame. */
export const DRAWN_CHART_TYPES: ReadonlySet<ChartSummaryType> = new Set([
	'bar',
	'column',
	'line',
	'area',
	'pie',
	'doughnut',
	'scatter',
	'radar',
]);

/** Theme fonts (latin typefaces) for chart text. */
export interface ChartThemeFonts {
	major?: string;
	minor?: string;
}

export interface RenderChartSpaceOptions extends ChartSvgOptions, ChartSummaryExtras {
	/** Size in CSS pixels. */
	width: number;
	height: number;
	/** The document theme's colour scheme; missing slots use the Office defaults. */
	theme?: ThemeColorScheme;
	fonts?: ChartThemeFonts;
}

/** A DrawingML colour scheme and fonts as the painter's theme palette. */
export function chartThemePalette(
	theme: ThemeColorScheme | undefined,
	fonts: ChartThemeFonts = {},
): ChartThemePalette {
	return {
		colors: THEME_SLOTS.map(
			(slot, index) => themeSlotHex(theme?.colors[slot]) ?? DEFAULT_THEME_COLORS[index] ?? '000000',
		),
		majorFont: fonts.major ?? 'Calibri Light',
		minorFont: fonts.minor ?? 'Calibri',
	};
}

/**
 * What the painter will show for a chart part: the family of its first chart group and whether
 * that family is drawn. A part with no chart group (a `cx:chartSpace`, an empty plot area) has no
 * family and is not drawn.
 */
export function chartSpaceKind(space: ChartSpace): {
	type?: ChartSummaryType;
	drawn: boolean;
} {
	if (!space.plotArea.groups.length) return { drawn: false };
	const type = chartSummaryFromSpace(space).chartType;
	return { type, drawn: DRAWN_CHART_TYPES.has(type) };
}

/** A standalone SVG document for a parsed chart part, from its cached values. */
export function renderChartSpaceSvg(space: ChartSpace, options: RenderChartSpaceOptions): string {
	const { width, height, theme, fonts, formatting, colorPalette, styleDefinition, ...svg } =
		options;
	const summary = chartSummaryFromSpace(space, {
		...(formatting ? { formatting } : {}),
		...(colorPalette !== undefined ? { colorPalette } : {}),
		...(styleDefinition ? { styleDefinition } : {}),
	});
	const view = chartSummaryView(summary, chartThemePalette(theme, fonts));
	return renderChartSvg(view, width, height, svg);
}
