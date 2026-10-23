// Series paint of a spreadsheet chart (its manual, legacy or palette colour) and the private
// colour style part (`colorsN.xml`) a Change Colors edit writes beside the chart part. The style
// part is not in `ChartSpace`, so it is written here, not by the chart writer.
import { chartPaletteSeriesColorChoice, findChartColorPalette } from '../../chart/color-palettes';
import {
	chartColorStyleId,
	chartColorStyleXml,
	CHART_COLOR_STYLE_CONTENT_TYPE,
	CHART_COLOR_STYLE_REL,
} from '../../chart/color-style';
import { drawingFillXml } from '../../drawingml/write-fill';
import type { DrawingColor, DrawingFill } from '../../drawingml/types';
import type { ChartObject, ChartSeries } from '../model';
import { chartDrawingColor } from '../edit/chart-series-fill';
import { RelationshipSet, type PackageWriter } from './package-writer';

/** The primary paint of a series: its imported fill when it can be written, else solid. */
export function chartSeriesDrawingFill(
	chart: ChartObject,
	series: ChartSeries,
	index: number,
): DrawingFill {
	if (series.fill && drawingFillXml(series.fill)) return series.fill;
	return { kind: 'solid', color: chartSeriesColor(chart, series, index) };
}

/** The series colour: manual, legacy, else the palette's choice for the series. */
function chartSeriesColor(chart: ChartObject, series: ChartSeries, index: number): DrawingColor {
	const palette = findChartColorPalette(chart.colorPalette ?? 10)!;
	return (
		series.drawingColor ??
		chartDrawingColor(series.color) ??
		chartPaletteSeriesColorChoice(palette ?? findChartColorPalette(10)!, index, chart.series.length)
	);
}

/** Writes a private color-style copy, so editing one chart cannot recolor another chart. */
export function writeChartColorStyle(
	writer: PackageWriter,
	chartPart: string,
	chart: ChartObject,
): void {
	if (chart.colorPalette === undefined) return;
	const palette = findChartColorPalette(chart.colorPalette);
	if (!palette) return;
	const source = writer.source;
	const existing = [...(source?.rels(chartPart).entries() ?? [])];
	const entry = existing.find(([, rel]) => rel.type.endsWith('/chartColorStyle'));
	const sourcePart = entry && source?.target(chartPart, entry[1]);
	const original = sourcePart && source?.text(sourcePart);
	if (original && chartColorStyleId(original) === palette.id) return;
	const part = writer.uniqueName(
		(n) => `xl/charts/colors${n}.xml`,
		source ? new Set(source.parts.keys()) : undefined,
	);
	writer.add(
		part,
		chartColorStyleXml(palette, original || undefined),
		CHART_COLOR_STYLE_CONTENT_TYPE,
	);
	const rels = new RelationshipSet();
	for (const [id, rel] of existing) rels.keep(id, rel);
	const target =
		chartPart.split('/').slice(0, -1).join('/') === 'xl/charts'
			? part.slice('xl/charts/'.length)
			: '/' + part;
	if (entry) rels.keep(entry[0], { ...entry[1], target });
	else rels.add(CHART_COLOR_STYLE_REL, target);
	writer.rels(chartPart, rels);
}
