// Chart parity harness: the neutral side. Projects `ChartSpace` onto the comparison shape
// (`chart-neutral-parity.ts`), applying each pptx normalisation explicitly (search "pptx
// normalisation") so the projection of a correctly parsed part equals the pptx projection.
import type {
	ChartAxis,
	ChartDataLabelOptions,
	ChartMarker,
	ChartPlotGroup,
	ChartSpace,
	ChartSpaceSeries,
	ChartTitle,
} from '../../../chart/index';
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
import {
	GROUP_TYPE,
	neutralCategories,
	numbers,
	runText,
} from './chart-neutral-parity-neutral-data';

function neutralLabels(
	options: ChartDataLabelOptions | undefined,
	leaderLines?: boolean,
): LabelShape | undefined {
	if (!options) return undefined;
	const shape = compact<LabelShape>({
		showValue: options.showValue,
		showCategory: options.showCategoryName,
		showSeriesName: options.showSeriesName,
		showPercent: options.showPercent,
		showLegendKey: options.showLegendKey,
		showBubbleSize: options.showBubbleSize,
		showLeaderLines: leaderLines,
		position: options.position,
		separator: options.separator,
		numberFormat: options.numberFormat?.formatCode.trim() || undefined,
	});
	return Object.keys(shape).length > 0 ? shape : undefined;
}

/**
 * pptx normalisation (`parseMarker`): a marker without a known `c:symbol` is `auto`, and a size
 * outside the schema range 2..72 is dropped.
 */
function neutralMarker(marker: ChartMarker | undefined): MarkerShape | undefined {
	if (!marker) return undefined;
	const symbols = [
		'circle',
		'dash',
		'diamond',
		'dot',
		'none',
		'picture',
		'plus',
		'square',
		'star',
		'triangle',
		'x',
		'auto',
	];
	const symbol = marker.symbol?.trim() ?? '';
	const size = marker.size;
	return compact<MarkerShape>({
		symbol: symbols.includes(symbol) ? symbol : 'auto',
		size:
			size !== undefined && Number.isInteger(size) && size >= 2 && size <= 72 ? size : undefined,
	});
}

function neutralSeries(
	series: ChartSpaceSeries,
	index: number,
	group: ChartPlotGroup,
	chartType: string,
	categories: string[],
): SeriesShape {
	const groupType = GROUP_TYPE[group.element] ?? 'unknown';
	const valueSource = series.values ?? series.yValues;
	let values = numbers(valueSource);
	// pptx normalisation: a series whose value cache is empty gets placeholder values 1+i, 2+i...
	// (one per category) so an unrendered chart still draws something (`buildChartSeries`).
	if (values.length === 0) values = categories.map((_, at) => at + 1 + index);
	// pptx normalisation: the name is trimmed, and a series with no name text is called "Series"
	// (`extractChartSeriesName`).
	const name = series.tx?.text?.trim() || 'Series';
	// The series format code is the data-label `c:numFmt` when written, else the value cache's
	// `c:formatCode` (both parsers' "linked to source" rule).
	const formatCode =
		series.dataLabels?.numberFormat?.formatCode.trim() ||
		(valueSource?.kind === 'numRef' ? valueSource.cache?.formatCode?.trim() : undefined) ||
		undefined;
	const labels = series.dataLabels;
	return compact<SeriesShape>({
		chartType: chartType === 'combo' ? groupType : chartType,
		idx: series.index,
		name,
		values,
		xValues: nonEmpty(numbers(series.xValues).map((value) => value ?? Number.NaN)),
		bubbleSizes: nonEmpty(numbers(series.bubbleSizes).map((value) => value ?? 0)),
		formatCode,
		smooth: series.smooth,
		invertIfNegative: series.invertIfNegative,
		explosion: series.explosion,
		// pptx normalisation: `c:shape` is read only inside a 3-D bar group.
		shape: groupType === 'bar3D' ? series.shape : undefined,
		marker: neutralMarker(series.marker),
		dataPoints: nonEmpty(
			series.dataPoints.flatMap((point) =>
				point.index === undefined
					? []
					: [
							compact({
								idx: point.index,
								explosion: point.explosion,
								invertIfNegative: point.invertIfNegative,
								bubble3D: point.bubble3D,
								marker: neutralMarker(point.marker),
							}),
						],
			),
		),
		labels: neutralLabels(labels, labels?.showLeaderLines),
		pointLabels: nonEmpty(
			(labels?.labels ?? []).flatMap((label) =>
				label.index === undefined
					? []
					: [
							compact<PointLabelShape>({
								idx: label.index,
								deleted: label.deleted,
								...neutralLabels(label),
								text: runText(label.tx?.rich) || undefined,
							}),
						],
			),
		),
	});
}

function neutralAxis(axis: ChartAxis): AxisShape {
	return compact<AxisShape>({
		kind: `${axis.kind}Ax`,
		id: axis.id,
		crossAxisId: axis.crossAxisId,
		position: axis.position,
		// Tolerance: the pptx model stores `deleted` only when true.
		deleted: axis.deleted === true ? true : undefined,
		crosses: axis.crosses,
		crossesAt: axis.crossesAt,
		crossBetween: axis.crossBetween,
		orientation: axis.scaling.orientation,
		min: axis.scaling.min,
		max: axis.scaling.max,
		logBase: axis.scaling.logBase,
		majorUnit: axis.majorUnit,
		minorUnit: axis.minorUnit,
		// Tolerance: the pptx model stores gridline flags only when present.
		majorGridlines: axis.majorGridlines || undefined,
		minorGridlines: axis.minorGridlines || undefined,
		majorTickMark: axis.majorTickMark,
		minorTickMark: axis.minorTickMark,
		tickLabelPosition: axis.tickLabelPosition,
		// Tolerance: pptx keeps `sourceLinked` as a plain boolean (absent reads as false), and drops
		// an empty format code.
		numberFormat: axis.numberFormat?.formatCode.trim()
			? {
					formatCode: axis.numberFormat.formatCode.trim(),
					sourceLinked: axis.numberFormat.sourceLinked === true,
				}
			: undefined,
		title: axisTitle(axis.title),
	});
}

function axisTitle(title: ChartTitle | undefined): string | undefined {
	const text = runText(title?.tx?.rich);
	return text ? text : undefined;
}

/**
 * pptx normalisation (`getChartDataForGraphicFrame`): the flat chart title is the FIRST text run
 * of the rich title (the full run list is `titleRuns`), else the linked cell's cached text, trimmed.
 */
function chartTitle(title: ChartTitle | undefined): string | undefined {
	const firstRun = title?.tx?.rich?.paragraphs.flatMap((paragraph) => paragraph.runs)[0];
	if (firstRun) return firstRun.text;
	const cached = title?.tx?.reference?.cache?.points
		.slice()
		.sort((left, right) => (left.index ?? 0) - (right.index ?? 0))
		.map((point) => point.value.trim())
		.find((value) => value.length > 0);
	return cached || undefined;
}

/** pptx normalisation: grouping is `stacked`, `percentStacked` or else `clustered` (+ standard). */
function grouping(raw: string | undefined): { grouping?: string; groupingStandard?: boolean } {
	if (raw === undefined) return {};
	const value = raw.trim();
	if (value === 'stacked' || value === 'percentStacked') return { grouping: value };
	return value === 'standard'
		? { grouping: 'clustered', groupingStandard: true }
		: { grouping: 'clustered' };
}

/** Projects the neutral model of one chart part. */
export function projectNeutralChart(space: ChartSpace): ChartParityShape {
	const groups = space.plotArea.groups;
	const first = groups[0];
	const elements = new Set(groups.map((group) => group.element));
	const chartType =
		elements.size >= 2 ? 'combo' : first ? (GROUP_TYPE[first.element] ?? 'unknown') : 'unknown';
	const { categories, categoryLevels } = neutralCategories(groups);
	// The placeholder values of an empty series count from the series' position in its own group.
	const series = groups.flatMap((group) =>
		group.series.map((entry, index) => ({ entry, group, index })),
	);
	const scatterStyles = ['none', 'line', 'lineMarker', 'marker', 'smooth', 'smoothMarker'];
	const legend = space.legend;
	const firstLabels = groups.find((group) => group.dataLabels)?.dataLabels;
	return compact<ChartParityShape>({
		chartType,
		...grouping(first?.grouping),
		// pptx normalisation: any `c:barDir` other than `bar` is a column chart.
		barDirection: first?.barDirection,
		gapWidth: first?.gapWidth,
		overlap: first?.overlap,
		gapDepth: first?.gapDepth,
		varyColors: first?.varyColors,
		firstSliceAngle: first?.firstSliceAngle,
		holeSize: first?.holeSize,
		scatterStyle:
			first?.scatterStyle && scatterStyles.includes(first.scatterStyle)
				? first.scatterStyle
				: undefined,
		// pptx normalisation: radar style, 3-D bar shape and wireframe are read only for their own
		// chart type, and the surface projection is told by the element name.
		radarStyle: chartType === 'radar' ? first?.radarStyle : undefined,
		barShape: chartType === 'bar3D' ? first?.shape : undefined,
		wireframe: chartType === 'surface' ? first?.wireframe : undefined,
		surfaceTopView: chartType === 'surface' ? first?.element === 'surfaceChart' : undefined,
		categories,
		categoryLevels,
		series: series.map(({ entry, group, index }) =>
			neutralSeries(entry, index, group, chartType, categories),
		),
		axes: space.plotArea.axes.map(neutralAxis).sort(byAxisId),
		title: chartTitle(space.title),
		// pptx normalisation (`resolveLegendOverlay`): a classic legend without `c:overlay` overlays.
		legend: legend
			? compact({ position: legend.position, overlay: legend.overlay ?? true })
			: undefined,
		groupLabels: neutralLabels(firstLabels, firstLabels?.showLeaderLines),
		plotVisibleOnly: space.plotVisibleOnly,
		dispBlanksAs: space.displayBlanksAs,
		autoTitleDeleted: space.autoTitleDeleted,
		style: space.style,
		view3D: space.view3D
			? (compact({ ...space.view3D }) as Record<string, number | boolean | undefined>)
			: undefined,
		date1904: space.date1904,
		roundedCorners: space.roundedCorners,
	});
}
