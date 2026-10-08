// Chart parity harness: the pptx side. Projects `PptxChartData` onto the comparison shape
// (`chart-neutral-parity.ts`); the pptx model is the reference, so this side is a plain mapping.
import type {
	PptxChartAxisFormatting,
	PptxChartData,
	PptxChartDataLabel,
	PptxChartDataLabelOptions,
	PptxChartMarker,
	PptxChartSeries,
	PptxChartType,
} from '../../core/types';
import {
	byAxisId,
	compact,
	nonEmpty,
	type AxisShape,
	type ChartParityShape,
	type LabelShape,
	type MarkerShape,
	type PointLabelShape,
	type SeriesShape,
} from './chart-neutral-parity';

function pptxLabels(options: PptxChartDataLabelOptions | undefined): LabelShape | undefined {
	if (!options) return undefined;
	const shape = compact<LabelShape>({
		showValue: options.showValue,
		showCategory: options.showCategory,
		showSeriesName: options.showSeriesName,
		showPercent: options.showPercent,
		showLegendKey: options.showLegendKey,
		showBubbleSize: options.showBubbleSize,
		showLeaderLines: options.showLeaderLines,
		position: options.position,
		separator: options.separator,
		numberFormat: options.numberFormat,
	});
	return Object.keys(shape).length > 0 ? shape : undefined;
}

function pptxPointLabel(label: PptxChartDataLabel): PointLabelShape {
	return compact<PointLabelShape>({
		idx: label.idx,
		deleted: label.deleted,
		showValue: label.showVal,
		showCategory: label.showCatName,
		showSeriesName: label.showSerName,
		showPercent: label.showPercent,
		showLegendKey: label.showLegendKey,
		showBubbleSize: label.showBubbleSize,
		showLeaderLines: label.showLeaderLines,
		position: label.position,
		separator: label.separator,
		numberFormat: label.numberFormat,
		// Cell-range text (`textFromCells`) comes from the c15 field table, which the neutral
		// model keeps only as raw `extLst`; only authored rich text is compared.
		text: label.textFromCells ? undefined : label.text,
	});
}

function pptxMarker(marker: PptxChartMarker | undefined): MarkerShape | undefined {
	return marker ? compact<MarkerShape>({ symbol: marker.symbol, size: marker.size }) : undefined;
}

function pptxSeries(series: PptxChartSeries, chartType: PptxChartType): SeriesShape {
	// The pptx model keeps blank points as 0 plus a parallel `blanks` mask; the comparison shape
	// carries them as null.
	const values = series.values.map((value, index) => (series.blanks?.[index] ? null : value));
	return compact<SeriesShape>({
		chartType: series.seriesChartType ?? chartType,
		idx: series.idx,
		name: series.name,
		values,
		xValues: nonEmpty(series.xValues),
		bubbleSizes: nonEmpty(series.bubbleSizes),
		formatCode: series.numberFormat,
		smooth: series.smooth,
		invertIfNegative: series.invertIfNegative,
		explosion: series.explosion,
		shape: series.shape,
		marker: pptxMarker(series.marker),
		dataPoints: nonEmpty(
			(series.dataPoints ?? []).map((point) =>
				compact({
					idx: point.idx,
					explosion: point.explosion,
					invertIfNegative: point.invertIfNegative,
					bubble3D: point.bubble3D,
					marker: pptxMarker(point.marker),
				}),
			),
		),
		labels: pptxLabels(series.dataLabelOptions),
		pointLabels: nonEmpty((series.dataLabels ?? []).map(pptxPointLabel)),
	});
}

function pptxAxis(axis: PptxChartAxisFormatting): AxisShape {
	return compact<AxisShape>({
		kind: axis.axisType,
		id: axis.axisId,
		crossAxisId: axis.crossAxisId,
		position: axis.axPos,
		deleted: axis.deleted,
		crosses: axis.crosses,
		crossesAt: axis.crossesAt,
		crossBetween: axis.crossBetween,
		orientation: axis.orientation,
		min: axis.min,
		max: axis.max,
		logBase: axis.logBase,
		majorUnit: axis.majorUnit,
		minorUnit: axis.minorUnit,
		majorGridlines: axis.majorGridlines,
		minorGridlines: axis.minorGridlines,
		majorTickMark: axis.majorTickMark,
		minorTickMark: axis.minorTickMark,
		tickLabelPosition: axis.tickLblPos,
		numberFormat: axis.numFmt
			? { formatCode: axis.numFmt.formatCode, sourceLinked: axis.numFmt.sourceLinked === true }
			: undefined,
		title: axis.titleText,
	});
}

/** Projects the pptx chart model of one chart part. */
export function projectPptxChart(data: PptxChartData): ChartParityShape {
	const style = data.style;
	return compact<ChartParityShape>({
		chartType: data.chartType,
		grouping: data.grouping,
		groupingStandard: data.groupingStandard,
		barDirection: data.barDirection,
		gapWidth: data.barGapWidth,
		overlap: data.barOverlap,
		gapDepth: data.gapDepth,
		varyColors: data.varyColors,
		firstSliceAngle: data.firstSliceAngle,
		holeSize: data.doughnutHoleSize,
		scatterStyle: data.scatterStyle,
		radarStyle: data.radarStyle,
		barShape: data.barShape,
		wireframe: data.wireframe,
		surfaceTopView: data.surfaceTopView,
		categories: data.categories,
		categoryLevels: data.categoryLevels,
		series: data.series.map((series) => pptxSeries(series, data.chartType)),
		axes: (data.axes ?? []).map(pptxAxis).sort(byAxisId),
		title: data.title,
		legend: style?.hasLegend
			? compact({ position: style.legendPosition, overlay: style.legendOverlay })
			: undefined,
		groupLabels: pptxLabels(style?.dataLabels),
		plotVisibleOnly: data.plotVisibleOnly,
		dispBlanksAs: data.chartChrome?.dispBlanksAs,
		autoTitleDeleted: data.chartChrome?.autoTitleDeleted,
		style: style?.styleId,
		view3D: data.view3D
			? compact({
					rotX: data.view3D.rotX,
					rotY: data.view3D.rotY,
					depthPercent: data.view3D.depthPercent,
					heightPercent: data.view3D.hPercent,
					rightAngleAxes: data.view3D.rAngAx,
					perspective: data.view3D.perspective,
				})
			: undefined,
		date1904: data.date1904,
		roundedCorners: data.roundedCorners,
	});
}
