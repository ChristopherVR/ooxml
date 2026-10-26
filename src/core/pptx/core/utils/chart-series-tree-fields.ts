import { EMU_PER_POINT } from '../../../units/index';
import type { ChartSpaceSeries } from '../../../chart/index';
import { chartColorFromNeutral, chartGradientFromNeutral } from './chart-from-neutral-paint';
import type { IPptxXmlLookupService } from '../services/PptxXmlLookupService';
import type { PptxChartSeries, PptxChartType, XmlObject } from '../types';
import { isLineDrawnChartType } from './chart-container-type-map';
import { parseChartDataLabelOptions, parseSeriesDataLabels } from './chart-data-label-parser';
import {
	parseChartDataPointPicture,
	parseImplicitBlipPictureFill,
} from './chart-datapoint-picture';
import type { ChartGradientCodec } from './chart-gradient-fill';
/**
 * Compatibility styling for a neutral chart series. Its solid and gradient paint, line width,
 * dash and no-fill fields come from neutral DrawingML. Source-tree readers remain for picture
 * relationships, c16 identity and the styled text of labels.
 *
 * @module chart-series-tree-fields
 */

import { neutralChartMarker, neutralChartPoints } from './chart-from-neutral-points';
import { parseChartUniqueId } from './chart-series-identity';

/** What the object-tree half of the chart adapter reads with: the runtime's services. */
export interface ChartTreeReader {
	xml: IPptxXmlLookupService;
	/** Resolves a colour node through the theme and the chart's colour map override. */
	parseColor: (node: XmlObject | undefined, placeholder?: string) => string | undefined;
	/** Resolves a theme typeface token (`+mn-lt`) to the deck's face. */
	resolveTypeface: (raw: string) => string;
	localName: (key: string) => string;
	colorStyleCodec?: ChartGradientCodec;
}

/** The series fields read from the object tree; the structural ones come from the neutral model. */
export type SeriesTreeFields = Pick<
	PptxChartSeries,
	| 'color'
	| 'gradientFill'
	| 'lineGradientFill'
	| 'dataPoints'
	| 'marker'
	| 'dataLabels'
	| 'dataLabelOptions'
	| 'picture'
	| 'impliedPicture'
	| 'lineNoFill'
	| 'lineWidth'
	| 'lineDashStyle'
	| 'uniqueId'
>;

/**
 * Reads the object-tree fields of one series. `containerChartType` is the series' own group type:
 * line-drawn series (line, scatter, radar, stock) author their colour on the outline.
 */
export function readSeriesTreeFields(
	seriesNode: XmlObject,
	containerChartType: PptxChartType | undefined,
	reader: ChartTreeReader,
	neutral: Pick<ChartSpaceSeries, 'spPr' | 'marker' | 'dataPoints'>,
): SeriesTreeFields {
	const { xml } = reader;
	const neutralShape = neutral.spPr ?? {};
	const colorAdapter = { parseColor: reader.parseColor };
	const color =
		(neutralShape.fill?.kind === 'solid'
			? chartColorFromNeutral(neutralShape.fill.color, reader)
			: undefined) ??
		(isLineDrawnChartType(containerChartType) && neutralShape.line?.fill?.kind === 'solid'
			? chartColorFromNeutral(neutralShape.line.fill.color, reader)
			: undefined);
	const dataPoints = neutralChartPoints(neutral, seriesNode, reader);
	const marker = neutralChartMarker(neutral.marker, reader);
	const dataLabels = parseSeriesDataLabels(seriesNode, xml, colorAdapter, reader.resolveTypeface);
	const labelGroup = xml.getChildByLocalName(seriesNode, 'dLbls');
	const dataLabelOptions = labelGroup
		? parseChartDataLabelOptions(labelGroup, xml, colorAdapter, reader.resolveTypeface)
		: undefined;
	const picture = parseChartDataPointPicture(seriesNode, xml);
	const impliedPicture = picture ? undefined : parseImplicitBlipPictureFill(seriesNode, xml);
	const lineNoFill = neutralShape.line?.fill?.kind === 'none';
	const lineWidthEmu = neutralShape.line?.widthEmu;
	const lineWidth =
		lineWidthEmu !== undefined && Number.isFinite(lineWidthEmu) && lineWidthEmu > 0
			? lineWidthEmu / EMU_PER_POINT
			: undefined;
	const dash = neutralShape.line?.dash;
	const uniqueId = parseChartUniqueId(seriesNode, reader.localName);
	const gradientFill = chartGradientFromNeutral(neutralShape.fill, reader);
	const lineGradientFill = isLineDrawnChartType(containerChartType)
		? chartGradientFromNeutral(neutralShape.line?.fill, reader)
		: undefined;
	return {
		color,
		...(gradientFill ? { gradientFill } : {}),
		...(lineGradientFill ? { lineGradientFill } : {}),
		...(dataPoints.length > 0 ? { dataPoints } : {}),
		...(marker ? { marker } : {}),
		...(dataLabels.length > 0 ? { dataLabels } : {}),
		...(dataLabelOptions ? { dataLabelOptions } : {}),
		...(picture ? { picture } : {}),
		...(impliedPicture ? { impliedPicture } : {}),
		...(lineNoFill ? { lineNoFill } : {}),
		...(lineWidth !== undefined ? { lineWidth } : {}),
		...(typeof dash === 'string' && dash.length > 0 ? { lineDashStyle: dash } : {}),
		...(uniqueId !== undefined ? { uniqueId } : {}),
	};
}
