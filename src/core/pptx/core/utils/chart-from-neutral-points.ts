import type { ChartMarker, ChartSpaceSeries } from '../../../chart/index';
import type { PptxChartDataPoint, PptxChartMarker, XmlObject } from '../types';
import { markerFromNeutral } from './chart-from-neutral-labels';
import { chartGradientFromNeutral, chartShapeFromNeutral } from './chart-from-neutral-paint';
import {
	parseChartDataPointPicture,
	parseImplicitBlipPictureFill,
} from './chart-datapoint-picture';
import { parseChartUniqueId } from './chart-series-identity';
import type { ChartTreeReader } from './chart-series-tree-fields';

export function neutralChartMarker(
	marker: ChartMarker | undefined,
	reader: ChartTreeReader,
): PptxChartMarker | undefined {
	if (!marker) return undefined;
	const result = markerFromNeutral(marker, undefined)!;
	const shape = chartShapeFromNeutral(marker.spPr, reader);
	if (shape) result.spPr = shape;
	return result;
}

/** Only picture host references and extension identities need the retained compatibility tree. */
export function neutralChartPoints(
	series: Pick<ChartSpaceSeries, 'dataPoints'>,
	source: XmlObject,
	reader: ChartTreeReader,
): PptxChartDataPoint[] {
	const nodes = [...reader.xml.getChildrenArrayByLocalName(source, 'dPt')];
	return series.dataPoints.flatMap((point) => {
		const idx = point.index;
		if (idx === undefined || !Number.isInteger(idx) || idx < 0 || idx > 0xffffffff) return [];
		const at = nodes.findIndex(
			(entry) => Number(reader.xml.getChildByLocalName(entry, 'idx')?.['@_val']) === idx,
		);
		const node = at >= 0 ? nodes.splice(at, 1)[0] : undefined;
		const result: PptxChartDataPoint = { idx };
		const shape = chartShapeFromNeutral(point.spPr, reader);
		if (shape) result.spPr = shape;
		const gradient = chartGradientFromNeutral(point.spPr?.fill, reader);
		if (gradient) result.gradientFill = gradient;
		const picture = node ? parseChartDataPointPicture(node, reader.xml) : undefined;
		if (picture) result.picture = picture;
		else {
			const implied = node ? parseImplicitBlipPictureFill(node, reader.xml) : undefined;
			if (implied) result.impliedPicture = implied;
		}
		if (
			point.explosion !== undefined &&
			Number.isInteger(point.explosion) &&
			point.explosion >= 0 &&
			point.explosion <= 0xffffffff
		)
			result.explosion = point.explosion;
		if (point.invertIfNegative !== undefined) result.invertIfNegative = point.invertIfNegative;
		const marker = neutralChartMarker(point.marker, reader);
		if (marker) result.marker = marker;
		if (point.bubble3D !== undefined) result.bubble3D = point.bubble3D;
		const uniqueId = node ? parseChartUniqueId(node, reader.localName) : undefined;
		if (uniqueId !== undefined) result.uniqueId = uniqueId;
		return [result];
	});
}
