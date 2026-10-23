// `cx:chartSpace` writer: chart data, text and titles, series with their data labels and layout
// properties, and axes. Elements and attributes follow the sequences of the chartex schema
// ([MS-ODRAWXML] 2.24.3).
import type {
	ChartExAxis,
	ChartExAxisScaling,
	ChartExChartData,
	ChartExData,
	ChartExDataLabelOptions,
	ChartExDataLabels,
	ChartExLayoutProperties,
	ChartExSeries,
	ChartExText,
	ChartExTitle,
} from './model';
import {
	attributes,
	cxElement,
	cxText,
	cxVal,
	formulaXml,
	numberFormatXml,
	raw,
	shapeXml,
	textBodyXml,
	type ChartWriteContext,
} from './write-util';

function dataXml(context: ChartWriteContext, data: ChartExData): string {
	const dimensions = data.dimensions
		.map((dimension) => {
			const levels = dimension.levels
				.map((level) =>
					cxElement(
						'lvl',
						level.points
							.map((point) => cxText('pt', point.value, attributes([['idx', point.index]])))
							.join('') + raw(context, level.extLst),
						attributes([
							['ptCount', level.pointCount],
							['formatCode', level.formatCode],
							['name', level.name],
						]),
					),
				)
				.join('');
			return cxElement(
				dimension.kind,
				formulaXml('f', dimension.formula) +
					formulaXml('nf', dimension.nameFormula) +
					levels +
					raw(context, dimension.extLst),
				attributes([['type', dimension.type]]),
			);
		})
		.join('');
	return cxElement('data', dimensions + raw(context, data.extLst), attributes([['id', data.id]]));
}

/** `cx:chartData`. */
export function chartDataXml(context: ChartWriteContext, chartData: ChartExChartData): string {
	const external = chartData.externalData;
	const externalXml = external
		? cxElement(
				'externalData',
				'',
				attributes([
					['r:id', external.relId],
					['cx:autoUpdate', external.autoUpdate],
				]),
			)
		: '';
	return cxElement(
		'chartData',
		externalXml + chartData.data.map((data) => dataXml(context, data)).join(''),
	);
}

/** `cx:tx`. */
export function textXml(context: ChartWriteContext, text: ChartExText | undefined): string {
	if (!text) return '';
	const data =
		text.formula !== undefined || text.value !== undefined
			? cxElement('txData', formulaXml('f', text.formula) + cxText('v', text.value))
			: '';
	return cxElement('tx', data + textBodyXml(context, text.rich, 'rich'));
}

/** A title (`cx:title`, `cx:unitsLabel` with `local`). */
export function titleXml(
	context: ChartWriteContext,
	title: ChartExTitle | undefined,
	local = 'title',
): string {
	if (!title) return '';
	return cxElement(
		local,
		textXml(context, title.tx) +
			shapeXml(context, title.spPr) +
			textBodyXml(context, title.txPr) +
			raw(context, title.extLst),
		attributes([
			['pos', title.position],
			['align', title.align],
			['overlay', title.overlay],
		]),
	);
}

function labelBody(context: ChartWriteContext, label: ChartExDataLabelOptions): string {
	const visibility = label.visibility;
	return (
		numberFormatXml(label.numberFormat) +
		shapeXml(context, label.spPr) +
		textBodyXml(context, label.txPr) +
		(visibility
			? cxElement(
					'visibility',
					'',
					attributes([
						['seriesName', visibility.seriesName],
						['categoryName', visibility.categoryName],
						['value', visibility.value],
					]),
				)
			: '') +
		cxText('separator', label.separator)
	);
}

function dataLabelsXml(context: ChartWriteContext, labels: ChartExDataLabels | undefined): string {
	if (!labels) return '';
	const points = labels.labels
		.map((label) =>
			cxElement(
				'dataLabel',
				labelBody(context, label) + raw(context, label.extLst),
				attributes([
					['idx', label.index],
					['pos', label.position],
				]),
			),
		)
		.join('');
	const hidden = labels.hidden
		.map((index) => cxElement('dataLabelHidden', '', attributes([['idx', index]])))
		.join('');
	return cxElement(
		'dataLabels',
		labelBody(context, labels) + points + hidden + raw(context, labels.extLst),
		attributes([['pos', labels.position]]),
	);
}

function layoutXml(context: ChartWriteContext, layout: ChartExLayoutProperties | undefined) {
	if (!layout) return '';
	const v = layout.visibility;
	const b = layout.binning;
	const g = layout.geography;
	return cxElement(
		'layoutPr',
		cxVal('parentLabelLayout', layout.parentLabelLayout) +
			cxVal('regionLabelLayout', layout.regionLabelLayout) +
			(v
				? cxElement(
						'visibility',
						'',
						attributes([
							['connectorLines', v.connectorLines],
							['meanLine', v.meanLine],
							['meanMarker', v.meanMarker],
							['nonoutliers', v.nonoutliers],
							['outliers', v.outliers],
						]),
					)
				: '') +
			(layout.aggregation ? cxElement('aggregation') : '') +
			(b
				? cxElement(
						'binning',
						cxVal('binSize', b.binSize) + cxVal('binCount', b.binCount),
						attributes([
							['intervalClosed', b.intervalClosed],
							['underflow', b.underflow],
							['overflow', b.overflow],
						]),
					)
				: '') +
			(g
				? cxElement(
						'geography',
						raw(context, g.geoCacheXml) + raw(context, g.extLst),
						attributes([
							['projectionType', g.projectionType],
							['viewedRegionType', g.viewedRegionType],
							['cultureLanguage', g.cultureLanguage],
							['cultureRegion', g.cultureRegion],
							['attribution', g.attribution],
						]),
					)
				: '') +
			(layout.statistics
				? cxElement(
						'statistics',
						'',
						attributes([['quartileMethod', layout.statistics.quartileMethod]]),
					)
				: '') +
			(layout.subtotals
				? cxElement('subtotals', layout.subtotals.map((index) => cxVal('idx', index)).join(''))
				: '') +
			raw(context, layout.extLst),
	);
}

/** `cx:series`. */
export function seriesXml(context: ChartWriteContext, series: ChartExSeries): string {
	const points = series.dataPoints
		.map((point) =>
			cxElement(
				'dataPt',
				shapeXml(context, point.spPr) + raw(context, point.extLst),
				attributes([['idx', point.index]]),
			),
		)
		.join('');
	return cxElement(
		'series',
		textXml(context, series.tx) +
			shapeXml(context, series.spPr) +
			textBodyXml(context, series.txPr) +
			raw(context, series.valueColorsXml) +
			raw(context, series.valueColorPositionsXml) +
			points +
			dataLabelsXml(context, series.dataLabels) +
			(series.data ? dataXml(context, series.data) : '') +
			cxVal('dataId', series.dataId) +
			layoutXml(context, series.layoutProperties) +
			series.axisIds.map((id) => cxVal('axisId', id)).join('') +
			raw(context, series.extLst),
		attributes([
			['layoutId', series.layoutId],
			['hidden', series.hidden],
			['ownerIdx', series.ownerIndex],
			['uniqueId', series.uniqueId],
			['formatIdx', series.formatIndex],
		]),
	);
}

function scalingXml(scaling: ChartExAxisScaling | undefined): string {
	if (!scaling) return '';
	if (scaling.kind === 'cat')
		return cxElement('catScaling', '', attributes([['gapWidth', scaling.gapWidth]]));
	return cxElement(
		'valScaling',
		'',
		attributes([
			['max', scaling.max],
			['min', scaling.min],
			['majorUnit', scaling.majorUnit],
			['minorUnit', scaling.minorUnit],
		]),
	);
}

const tickMarksXml = (local: string, marks: { type?: string } | undefined) =>
	marks ? cxElement(local, '', attributes([['type', marks.type]])) : '';

/** `cx:axis`. */
export function axisXml(context: ChartWriteContext, axis: ChartExAxis): string {
	const units = axis.units;
	const gridlines = (local: string, value: ChartExAxis['majorGridlines']) =>
		value ? cxElement(local, shapeXml(context, value.spPr)) : '';
	return cxElement(
		'axis',
		scalingXml(axis.scaling) +
			titleXml(context, axis.title) +
			(units
				? cxElement(
						'units',
						titleXml(context, units.label, 'unitsLabel') + raw(context, units.extLst),
						attributes([['unit', units.unit]]),
					)
				: '') +
			gridlines('majorGridlines', axis.majorGridlines) +
			gridlines('minorGridlines', axis.minorGridlines) +
			tickMarksXml('majorTickMarks', axis.majorTickMarks) +
			tickMarksXml('minorTickMarks', axis.minorTickMarks) +
			(axis.tickLabels ? cxElement('tickLabels', raw(context, axis.tickLabels.extLst)) : '') +
			numberFormatXml(axis.numberFormat) +
			shapeXml(context, axis.spPr) +
			textBodyXml(context, axis.txPr) +
			raw(context, axis.extLst),
		attributes([
			['id', axis.id],
			['hidden', axis.hidden],
		]),
	);
}
