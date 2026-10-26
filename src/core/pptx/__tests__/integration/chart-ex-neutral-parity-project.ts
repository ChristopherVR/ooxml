// ChartEx parity harness (chart-ex-neutral-parity.test.ts): one comparison shape for the fields the
// pptx ChartEx reader (`PptxChartData`) and the neutral chartex model (`ChartExSpace`) share, and a
// projection from each. The pptx normalisations (trimmed text, the `Series N` default name, `[0]`
// for a series with no values, numbers through `parseFloat`, layout options only for their family)
// are applied explicitly on the neutral side.
import type {
	ChartExAutoNumber,
	ChartExData,
	ChartExDataLabel,
	ChartExDimension,
	ChartExSeries,
	ChartExSpace,
	ChartExText,
} from '../../../chart/index';
import type { PptxChartData } from '../../index';

export interface ChartExShape {
	chartType: string;
	title?: string;
	legend?: { position?: string };
	categories: string[];
	categoryLevels?: string[][];
	hasDataLabels: boolean;
	series: Record<string, unknown>[];
	axes: Record<string, unknown>[];
}

/** Drops `undefined` entries (pptx writes some optional keys as explicit `undefined`). */
export function compact<T>(value: T): T {
	if (Array.isArray(value)) return value.map(compact) as T;
	if (!value || typeof value !== 'object') return value;
	return Object.fromEntries(
		Object.entries(value)
			.filter(([, item]) => item !== undefined)
			.map(([key, item]) => [key, compact(item)]),
	) as T;
}

const nonEmpty = <T extends object>(value: T): T | undefined =>
	Object.values(value).some((item) => item !== undefined) ? value : undefined;

// --- pptx side -------------------------------------------------------------------------------

/** The shared fields of a ChartEx `PptxChartData`. */
export function projectPptxChartEx(chart: PptxChartData): ChartExShape {
	return compact({
		chartType: chart.chartType,
		title: chart.title,
		legend: chart.style?.hasLegend ? { position: chart.style.legendPosition } : undefined,
		categories: chart.categories,
		categoryLevels: chart.categoryLevels,
		hasDataLabels: chart.style?.hasDataLabels === true,
		series: chart.series.map((series) => ({
			name: series.name,
			values: series.values,
			color: series.color?.toUpperCase(),
			labelOptions: series.dataLabelOptions
				? nonEmpty({
						position: series.dataLabelOptions.position,
						numberFormat: series.dataLabelOptions.numberFormat,
						showValue: series.dataLabelOptions.showValue,
						showCategoryName: series.dataLabelOptions.showCategory,
						showSeriesName: series.dataLabelOptions.showSeriesName,
					})
				: undefined,
			labels: series.dataLabels?.map((label) => ({
				index: label.idx,
				deleted: label.deleted === true ? true : undefined,
				position: label.position,
				numberFormat: label.numberFormat,
				showValue: label.showVal,
				showCategoryName: label.showCatName,
				showSeriesName: label.showSerName,
			})),
			histogram: series.histogramOptions,
			waterfall: series.waterfallOptions,
			boxWhisker: series.boxWhiskerOptions ? nonEmpty(series.boxWhiskerOptions) : undefined,
			treemap: series.treemapOptions,
			regionMap: series.regionMapOptions
				? nonEmpty({
						regionLabelLayout: series.regionMapOptions.regionLabelLayout,
						projectionType: series.regionMapOptions.projectionType,
						viewedRegionType: series.regionMapOptions.viewedRegionType,
						cultureLanguage: series.regionMapOptions.cultureLanguage,
						cultureRegion: series.regionMapOptions.cultureRegion,
						attribution: series.regionMapOptions.attribution,
						entityIds: series.regionMapOptions.entityIds,
						categorySourceIndices: series.regionMapOptions.categorySourceIndices,
						valueSourceIndices: series.regionMapOptions.valueSourceIndices,
						geoCache: series.regionMapOptions.geographyCache ? true : undefined,
					})
				: undefined,
		})),
		axes: (chart.axes ?? []).map((axis) => ({
			id: axis.axisId,
			kind: axis.axisType === 'catAx' ? 'cat' : 'val',
			hidden: axis.deleted === true,
			title: axis.titleText,
			displayUnits: axis.displayUnitsValue,
			displayUnitsLabel:
				typeof axis.displayUnitsLabel === 'string'
					? axis.displayUnitsLabel
					: axis.displayUnitsLabel?.text,
			numberFormat: axis.numFmt,
			majorGridlines: axis.majorGridlines === true,
			minorGridlines: axis.minorGridlines === true,
			tickLabels: axis.tickLblPos !== 'none',
			min: axis.min,
			max: axis.max,
			majorUnit: axis.majorUnit,
			minorUnit: axis.minorUnit,
		})),
	});
}

// --- neutral side ----------------------------------------------------------------------------

const LAYOUT_TYPES: Record<string, string> = {
	waterfall: 'waterfall',
	funnel: 'funnel',
	treemap: 'treemap',
	sunburst: 'sunburst',
	boxWhisker: 'boxWhisker',
	clusteredColumn: 'histogram',
	paretoLine: 'histogram',
	regionMap: 'regionMap',
};

/** The rich text runs joined, else the cached value, trimmed (`resolveCxTitleText`). */
function textOf(tx: ChartExText | undefined): string | undefined {
	const runs = tx?.rich?.paragraphs.flatMap((p) => p.runs.filter((r) => r.text !== '\n')) ?? [];
	if (runs.length > 0) return runs.map((run) => run.text).join('');
	return tx?.value?.trim() || undefined;
}

function seriesData(space: ChartExSpace, series: ChartExSeries): ChartExData | undefined {
	return (
		series.data ??
		(series.dataId === undefined
			? undefined
			: space.chartData.data.find((data) => data.id === series.dataId))
	);
}

const dimensionOf = (
	data: ChartExData | undefined,
	kind: ChartExDimension['kind'],
	type: string,
	fallbackType: string,
) => {
	const all = data?.dimensions.filter((d) => d.kind === kind) ?? [];
	return all.find((d) => (d.type?.trim() || fallbackType) === type) ?? all[0];
};

const levelText = (dimension: ChartExDimension | undefined) =>
	dimension?.levels.map((level) => level.points.map((point) => point.value.trim())) ?? [];

const parseNumber = (value: ChartExAutoNumber | undefined) => value;

/** The family options pptx reads for a series (`chart-cx-*.ts`). */
function familyOptions(space: ChartExSpace, series: ChartExSeries) {
	const layout = series.layoutProperties;
	const out: Record<string, unknown> = {};
	if (series.layoutId === 'clusteredColumn' || series.layoutId === 'paretoLine') {
		const binning = layout?.binning;
		out.histogram = {
			layout: series.layoutId === 'paretoLine' ? 'pareto' : 'histogram',
			binSize: binning?.binSize,
			binCount:
				binning?.binCount !== undefined && Number.isSafeInteger(binning.binCount)
					? binning.binCount
					: undefined,
			intervalClosed:
				binning?.intervalClosed === 'l' || binning?.intervalClosed === 'r'
					? binning.intervalClosed
					: undefined,
			underflow: parseNumber(binning?.underflow),
			overflow: parseNumber(binning?.overflow),
			aggregateByCategory: layout?.aggregation ? true : undefined,
		};
	}
	if (series.layoutId === 'waterfall' && layout) {
		const connectorLines = layout.visibility?.connectorLines;
		if (connectorLines !== undefined || layout.subtotals)
			out.waterfall = { connectorLines, subtotalIndices: layout.subtotals };
	}
	if (series.layoutId === 'boxWhisker') {
		const visibility = layout?.visibility;
		const method = layout?.statistics?.quartileMethod;
		out.boxWhisker = nonEmpty({
			quartileMethod: method === 'inclusive' || method === 'exclusive' ? method : undefined,
			showMeanLine: visibility?.meanLine,
			showMeanMarker: visibility?.meanMarker,
			showInnerPoints: visibility?.nonoutliers,
			showOutlierPoints: visibility?.outliers,
		});
	}
	const parent = layout?.parentLabelLayout?.trim();
	if (parent === 'none' || parent === 'banner' || parent === 'overlapping')
		out.treemap = { parentLabelLayout: parent };
	if (series.layoutId === 'regionMap') {
		const data = seriesData(space, series);
		const indexed = (kind: ChartExDimension['kind'], type: string) => {
			const level = data?.dimensions.find((d) => d.kind === kind && d.type === type)?.levels[0];
			return {
				values: level?.points.map((point) => point.value.trim()) ?? [],
				indices: level?.points.map((point, position) => point.index ?? position) ?? [],
			};
		};
		const geography = layout?.geography;
		const entityIds = indexed('strDim', 'entityId');
		const categories = indexed('strDim', 'cat');
		const values = indexed('numDim', 'colorVal');
		const pick = (value: string | undefined, allowed: string[]) =>
			value !== undefined && allowed.includes(value) ? value : undefined;
		out.regionMap = nonEmpty({
			regionLabelLayout: pick(layout?.regionLabelLayout, ['none', 'bestFitOnly', 'showAll']),
			projectionType: pick(geography?.projectionType, ['mercator', 'miller', 'robinson', 'albers']),
			viewedRegionType: pick(geography?.viewedRegionType, [
				'dataOnly',
				'postalCode',
				'county',
				'state',
				'countryRegion',
				'countryRegionList',
				'world',
			]),
			cultureLanguage: geography?.cultureLanguage,
			cultureRegion: geography?.cultureRegion,
			attribution: geography?.attribution,
			entityIds: entityIds.values.length ? entityIds.values : undefined,
			categorySourceIndices: categories.indices.length ? categories.indices : undefined,
			valueSourceIndices: values.indices.length ? values.indices : undefined,
			geoCache: geography?.geoCacheXml ? true : undefined,
		});
	}
	return out;
}

const LABEL_POSITIONS = ['bestFit', 'b', 'ctr', 'inBase', 'inEnd', 'l', 'outEnd', 'r', 't'];
const labelPosition = (value: string | undefined) =>
	value !== undefined && LABEL_POSITIONS.includes(value) ? value : undefined;

function labels(series: ChartExSeries) {
	const group = series.dataLabels;
	if (!group) return {};
	const visibility = group.visibility;
	return {
		labelOptions: nonEmpty({
			position: labelPosition(group.position),
			numberFormat: group.numberFormat?.formatCode.trim() || undefined,
			showValue: visibility?.value,
			showCategoryName: visibility?.categoryName,
			showSeriesName: visibility?.seriesName,
		}),
		labels: [
			...group.labels,
			...group.hidden
				.filter((index) => !group.labels.some((label) => label.index === index))
				.map((index): ChartExDataLabel => ({ index })),
		].map((label) => {
			const own = label.visibility ?? visibility;
			return {
				index: label.index ?? 0,
				deleted: group.hidden.includes(label.index ?? 0) ? true : undefined,
				position: labelPosition(label.position),
				numberFormat: label.numberFormat?.formatCode.trim() || undefined,
				showValue: own ? own.value === true : undefined,
				showCategoryName: own ? own.categoryName === true : undefined,
				showSeriesName: own ? own.seriesName === true : undefined,
			};
		}),
	};
}

const UNITS: Record<string, number> = {
	hundreds: 100,
	thousands: 1_000,
	tenThousands: 10_000,
	hundredThousands: 100_000,
	millions: 1_000_000,
	tenMillions: 10_000_000,
	hundredMillions: 100_000_000,
	billions: 1_000_000_000,
	trillions: 1_000_000_000_000,
};

/** The shared fields of a neutral chartex model. */
export function projectNeutralChartEx(space: ChartExSpace): ChartExShape {
	const all = space.plotArea.series;
	const firstData = all[0] ? seriesData(space, all[0]) : undefined;
	const categoryDimension = dimensionOf(firstData, 'strDim', 'cat', 'cat');
	const levels = levelText(categoryDimension);
	const anyLabels = all.some((series) => {
		const v = series.dataLabels?.visibility;
		return v?.value === true || v?.categoryName === true || v?.seriesName === true;
	});
	return compact({
		chartType: LAYOUT_TYPES[all[0]?.layoutId ?? ''] ?? 'funnel',
		title: textOf(space.title?.tx),
		legend: space.legend ? { position: space.legend.position } : undefined,
		categories: (levels[0] ?? []).filter((value) => value),
		categoryLevels: categoryDimension ? levels : undefined,
		hasDataLabels: anyLabels,
		series: all.map((series, index) => {
			const values = levelText(dimensionOf(seriesData(space, series), 'numDim', 'val', 'val'))[0]
				?.map((value) => Number.parseFloat(value))
				.filter((value) => Number.isFinite(value));
			const fill = series.spPr?.fill;
			return {
				name: series.tx?.value?.trim() || `Series ${index + 1}`,
				values: values?.length ? values : [0],
				color:
					fill?.kind === 'solid' && fill.color.kind === 'srgb'
						? `#${fill.color.value.toUpperCase()}`
						: undefined,
				...labels(series),
				...familyOptions(space, series),
			};
		}),
		axes: space.plotArea.axes.map((axis) => ({
			id: axis.id,
			kind: axis.scaling?.kind === 'cat' ? 'cat' : 'val',
			hidden: axis.hidden === true,
			title: textOf(axis.title?.tx),
			displayUnits: axis.units?.unit === undefined ? undefined : UNITS[axis.units.unit],
			displayUnitsLabel:
				axis.units?.unit !== undefined && UNITS[axis.units.unit]
					? textOf(axis.units.label?.tx)
					: undefined,
			numberFormat: axis.numberFormat?.formatCode.trim()
				? {
						formatCode: axis.numberFormat.formatCode.trim(),
						sourceLinked: axis.numberFormat.sourceLinked === true,
					}
				: undefined,
			majorGridlines: axis.majorGridlines !== undefined,
			minorGridlines: axis.minorGridlines !== undefined,
			tickLabels: axis.tickLabels !== undefined,
			min:
				axis.scaling?.kind === 'val' && typeof axis.scaling.min === 'number'
					? axis.scaling.min
					: undefined,
			max:
				axis.scaling?.kind === 'val' && typeof axis.scaling.max === 'number'
					? axis.scaling.max
					: undefined,
			majorUnit:
				axis.scaling?.kind === 'val' && typeof axis.scaling.majorUnit === 'number'
					? axis.scaling.majorUnit
					: undefined,
			minorUnit:
				axis.scaling?.kind === 'val' && typeof axis.scaling.minorUnit === 'number'
					? axis.scaling.minorUnit
					: undefined,
		})),
	});
}
