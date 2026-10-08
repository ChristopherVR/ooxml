// `c:chartSpace` writer: series (`c:ser`), markers and data points. One element order serves every
// series type (bar, line, scatter, pie, area, radar, bubble, surface), because the model only holds
// what the type allows: idx, order, tx, spPr, invertIfNegative, marker, explosion, dPt, dLbls, cat,
// val, xVal, yVal, bubbleSize, smooth, shape, bubble3D, extLst. Trendlines and error bars are not
// modelled, so they are not written.
import type { ChartDataPoint, ChartMarker, ChartSpaceSeries } from './model-series';
import { shapePropertiesXml } from './write-shape';
import { chartTextXml, dataLabelsXml, dataSourceXml } from './write-text';
import { elementXml, raw, valXml, type ChartWriteContext } from './write-util';

/** A marker (`c:marker`). */
export function markerXml(context: ChartWriteContext, marker: ChartMarker | undefined): string {
	if (!marker) return '';
	return elementXml(
		'marker',
		valXml('symbol', marker.symbol) +
			valXml('size', marker.size) +
			shapePropertiesXml(context, marker.spPr) +
			raw(context, marker.extLst),
	);
}

function dataPointXml(context: ChartWriteContext, point: ChartDataPoint): string {
	return elementXml(
		'dPt',
		valXml('idx', point.index) +
			valXml('invertIfNegative', point.invertIfNegative) +
			markerXml(context, point.marker) +
			valXml('bubble3D', point.bubble3D) +
			valXml('explosion', point.explosion) +
			shapePropertiesXml(context, point.spPr) +
			raw(context, point.extLst),
	);
}

/** One series; `position` stands in for a missing index or order (both are required). */
export function seriesXml(
	context: ChartWriteContext,
	series: ChartSpaceSeries,
	position: number,
): string {
	return elementXml(
		'ser',
		valXml('idx', series.index ?? position) +
			valXml('order', series.order ?? position) +
			chartTextXml(context, series.tx, true) +
			shapePropertiesXml(context, series.spPr) +
			valXml('invertIfNegative', series.invertIfNegative) +
			markerXml(context, series.marker) +
			valXml('explosion', series.explosion) +
			series.dataPoints.map((point) => dataPointXml(context, point)).join('') +
			dataLabelsXml(context, series.dataLabels) +
			dataSourceXml('cat', series.categories) +
			dataSourceXml('val', series.values) +
			dataSourceXml('xVal', series.xValues) +
			dataSourceXml('yVal', series.yValues) +
			dataSourceXml('bubbleSize', series.bubbleSizes) +
			valXml('smooth', series.smooth) +
			valXml('shape', series.shape) +
			valXml('bubble3D', series.bubble3D) +
			raw(context, series.extLst),
	);
}
