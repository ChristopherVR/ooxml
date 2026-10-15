import {
	computeHierarchicalSunburstArcs as coreHierarchicalArcs,
	computeSunburstArcs as coreArcs,
} from 'ooxml-core/chart';
import type { SunburstArc as CoreSunburstArc } from 'ooxml-core/chart';
import type { PptxChartData } from 'ooxml-core/pptx';

import { paletteColor } from './chart-view-model';

/**
 * One sunburst ring arc segment: the shared geometry (`ooxml-core/chart`, `sunburst-arcs`) with
 * its palette index resolved to this chart's fill colour.
 */
export interface SunburstArc extends Omit<CoreSunburstArc, 'colorIndex'> {
	/** Fill colour. */
	fill: string;
}

const withFill = (
	arcs: readonly CoreSunburstArc[],
	colorPalette: readonly string[] | undefined,
): SunburstArc[] =>
	arcs.map(({ colorIndex, ...arc }) => ({ ...arc, fill: paletteColor(colorIndex, colorPalette) }));

/** Compute the legacy one-ring-per-series sunburst geometry. */
export function computeSunburstArcs(
	series: PptxChartData['series'],
	cx: number,
	cy: number,
	maxR: number,
	colorPalette: readonly string[] | undefined,
): SunburstArc[] {
	return withFill(coreArcs(series, cx, cy, maxR), colorPalette);
}

/**
 * Compute one aligned ring per ChartEx category level. Levels are stored leaf first, so the
 * highest parent level becomes the innermost ring.
 */
export function computeHierarchicalSunburstArcs(
	categoryLevels: ReadonlyArray<ReadonlyArray<string>>,
	values: ReadonlyArray<number>,
	cx: number,
	cy: number,
	maxR: number,
	colorPalette: readonly string[] | undefined,
): SunburstArc[] {
	return withFill(coreHierarchicalArcs(categoryLevels, values, cx, cy, maxR), colorPalette);
}
