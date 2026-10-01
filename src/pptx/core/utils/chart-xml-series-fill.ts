/**
 * Series fill (`c:ser/c:spPr`) for SDK-generated ChartML. Split out of
 * `chart-xml-generator.ts` to keep that file within the repo's size limit.
 *
 * @module utils/chart-xml-series-fill
 */
import type { PptxChartSeries, XmlObject } from '../types';
import { buildChartGradFillXml } from './chart-gradient-fill-writer';
import type { ChartFamily } from './chart-xml-container-map';

/** Families whose series colour is authored on the outline (`a:ln`). */
const OUTLINE_COLOR_FAMILIES = new Set<ChartFamily>(['line', 'radar', 'scatter']);

/**
 * Whether a generated series of `family` is drawn as a line with no fillable
 * area, so a gradient fill is ignored (see `chart-gradient-fill-writer.ts`).
 */
export function isLineDrawnChartFamily(family: ChartFamily): boolean {
	return OUTLINE_COLOR_FAMILIES.has(family) || family === 'stock';
}

/**
 * The `c:spPr` for a generated series: its gradient when it has one on an
 * area-filled family, otherwise its solid colour (on `a:ln` for line-drawn
 * families), or `undefined` for an automatic fill.
 */
export function buildGeneratedSeriesSpPr(
	series: PptxChartSeries,
	family: ChartFamily,
): XmlObject | undefined {
	if (series.gradientFill && !isLineDrawnChartFamily(family)) {
		return { 'a:gradFill': buildChartGradFillXml(series.gradientFill) };
	}
	if (!series.color) {
		return undefined;
	}
	const fill = {
		'a:solidFill': { 'a:srgbClr': { '@_val': series.color.replace(/^#/u, '').toUpperCase() } },
	};
	return OUTLINE_COLOR_FAMILIES.has(family) ? { 'a:ln': fill } : fill;
}
