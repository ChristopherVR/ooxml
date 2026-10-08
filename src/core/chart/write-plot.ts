// `c:chartSpace` writer: the plot area and its chart groups. Every group type is written in one
// element order that agrees with each type's sequence in ECMA-376 Part 1, 21.2 (the model holds
// only what the type allows): ofPieType, scatterStyle, radarStyle, wireframe, barDir, grouping,
// varyColors, ser, dLbls, gapWidth, overlap, dropLines, hiLowLines, upDownBars, marker, smooth,
// gapDepth, shape, firstSliceAng, holeSize, bubble3D, bubbleScale, showNegBubbles, sizeRepresents,
// serLines, axId, extLst. Groups come before axes, as Office writes them.
import type { ChartPlotArea, ChartPlotGroup } from './model';
import type { ChartLines, ChartShapeProperties } from './model-series';
import { CHART_GROUP_ELEMENTS } from './parse-plot';
import { axisXml } from './write-axes';
import { seriesXml } from './write-series';
import { layoutXml, shapePropertiesXml } from './write-shape';
import { dataLabelsXml } from './write-text';
import { elementXml, raw, valXml, type ChartWriteContext } from './write-util';

const linesXml = (
	context: ChartWriteContext,
	local: string,
	present: boolean | undefined,
	shape: ChartShapeProperties | undefined,
) => (present ? elementXml(local, shapePropertiesXml(context, shape)) : '');

const barsXml = (context: ChartWriteContext, local: string, bars: ChartLines | undefined) =>
	bars ? elementXml(local, shapePropertiesXml(context, bars.spPr)) : '';

function upDownBarsXml(context: ChartWriteContext, group: ChartPlotGroup): string {
	if (!group.upDownBars) return '';
	return elementXml(
		'upDownBars',
		valXml('gapWidth', group.upDownBarsGapWidth) +
			barsXml(context, 'upBars', group.upBars) +
			barsXml(context, 'downBars', group.downBars),
	);
}

/** The element of a group: its own when it is a chart group element, else from kind and 3-D. */
function groupElement(group: ChartPlotGroup): string {
	if (Object.hasOwn(CHART_GROUP_ELEMENTS, group.element)) return group.element;
	const match = Object.entries(CHART_GROUP_ELEMENTS).find(
		([, value]) => value.kind === group.kind && value.is3D === group.is3D,
	);
	if (!match) throw new Error(`No chart group element for ${group.kind}`);
	return match[0];
}

/** One chart group (`c:barChart`, `c:pie3DChart`...). */
export function groupXml(context: ChartWriteContext, group: ChartPlotGroup): string {
	return elementXml(
		groupElement(group),
		valXml('ofPieType', group.ofPieType) +
			valXml('scatterStyle', group.scatterStyle) +
			valXml('radarStyle', group.radarStyle) +
			valXml('wireframe', group.wireframe) +
			valXml('barDir', group.barDirection) +
			valXml('grouping', group.grouping) +
			valXml('varyColors', group.varyColors) +
			group.series.map((series, position) => seriesXml(context, series, position)).join('') +
			dataLabelsXml(context, group.dataLabels) +
			valXml('gapWidth', group.gapWidth) +
			valXml('overlap', group.overlap) +
			linesXml(context, 'dropLines', group.dropLines, group.dropLinesSpPr) +
			linesXml(context, 'hiLowLines', group.hiLowLines, group.hiLowLinesSpPr) +
			upDownBarsXml(context, group) +
			valXml('marker', group.marker) +
			valXml('smooth', group.smooth) +
			valXml('gapDepth', group.gapDepth) +
			valXml('shape', group.shape) +
			valXml('firstSliceAng', group.firstSliceAngle) +
			valXml('holeSize', group.holeSize) +
			valXml('bubble3D', group.bubble3D) +
			valXml('bubbleScale', group.bubbleScale) +
			valXml('showNegBubbles', group.showNegativeBubbles) +
			valXml('sizeRepresents', group.sizeRepresents) +
			linesXml(context, 'serLines', group.seriesLines, group.seriesLinesSpPr) +
			group.axisIds.map((id) => valXml('axId', id)).join('') +
			raw(context, group.extLst),
	);
}

/** The plot area (`c:plotArea`). */
export function plotAreaXml(context: ChartWriteContext, plotArea: ChartPlotArea): string {
	return elementXml(
		'plotArea',
		layoutXml(context, plotArea.layout) +
			plotArea.groups.map((group) => groupXml(context, group)).join('') +
			plotArea.axes.map((axis) => axisXml(context, axis)).join('') +
			shapePropertiesXml(context, plotArea.spPr) +
			raw(context, plotArea.extLst),
	);
}
