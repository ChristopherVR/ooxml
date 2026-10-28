import { chartExSeriesPaint } from './summary-ex-series';
// Builds neutral SVG summaries for supported Office 2016 extended chart parts: waterfall,
// histogram, treemap, sunburst, funnel and box-and-whisker. Word uses cached points; Excel can
// resolve the retained formulas from live cells. Region maps and Pareto remain unsummarised.
import type { DrawingColor } from '../../drawingml/types';
import { parseXml, type XmlElement } from '../../xml/index';
import type { ChartExData, ChartExDimension, ChartExSeries, ChartExSpace } from '../ex/model';
import { isChartExRoot, parseChartExSpace } from '../ex/parse';
import type { ChartSummary, ChartSummarySeries, HistogramSummaryOptions } from './summary';

/** The data of a series: its own `cx:data`, else the shared one named by `cx:dataId`. */
function seriesData(space: ChartExSpace, series: ChartExSeries): ChartExData | undefined {
	if (series.data) return series.data;
	const shared = space.chartData.data;
	return series.dataId === undefined ? shared[0] : shared.find((data) => data.id === series.dataId);
}

interface Dimension {
	formula?: string;
	points: string[];
}

const findDimension = (
	data: ChartExData | undefined,
	kind: ChartExDimension['kind'],
	type: string,
) => data?.dimensions.find((d) => d.kind === kind && (d.type ?? '') === type);

/** The points of one level by index, trimmed, holes as `''`. */
function levelPoints(level: ChartExDimension['levels'][number]): string[] {
	const points: string[] = [];
	for (const point of level.points) points[point.index ?? points.length] = point.value.trim();
	for (let i = 0; i < Math.max(level.pointCount ?? 0, points.length); i++) points[i] ??= '';
	return points;
}

/** One dimension's formula and its first level of cached points, holes as `''`. */
function dimension(data: ChartExData | undefined, kind: ChartExDimension['kind'], type: string) {
	const element = findDimension(data, kind, type);
	if (!element) return undefined;
	const level = element.levels[0];
	const out: Dimension = { points: level ? levelPoints(level) : [] };
	const formula = element.formula?.formula.trim();
	if (formula) out.formula = formula;
	return out;
}

/** Every level of the category dimension, leaf first, holes as `''` (treemap and sunburst). */
const categoryLevels = (data: ChartExData | undefined): string[][] =>
	(findDimension(data, 'strDim', 'cat')?.levels ?? []).map(levelPoints);

/** The title text: `cx:v` of `cx:txData`, or the runs of `cx:rich` (line breaks left out). */
function titleOf(space: ChartExSpace): string | undefined {
	const tx = space.title?.tx;
	const data = tx?.value?.trim();
	if (data) return data;
	if (!tx?.rich) return undefined;
	const runs = tx.rich.paragraphs.flatMap((paragraph) =>
		paragraph.runs.filter((run) => run.text !== '\n').map((run) => run.text),
	);
	return runs.join('') || undefined;
}

function seriesFill(series: ChartExSeries): Pick<ChartSummarySeries, 'drawingColor'> {
	const fill = series.spPr?.fill;
	const color: DrawingColor | undefined = fill?.kind === 'solid' ? fill.color : undefined;
	return color ? { drawingColor: color } : {};
}

function histogramOptions(series: ChartExSeries): HistogramSummaryOptions {
	const layout = series.layoutProperties;
	const binning = layout?.binning;
	const out: HistogramSummaryOptions = {};
	if (binning?.binSize !== undefined) out.binSize = binning.binSize;
	if (binning?.binCount !== undefined && binning.binCount >= 0)
		out.binCount = Math.trunc(binning.binCount);
	const closed = binning?.intervalClosed;
	if (closed === 'l' || closed === 'r') out.intervalClosed = closed;
	for (const key of ['underflow', 'overflow'] as const) {
		const value = binning?.[key];
		if (value !== undefined) out[key] = value;
	}
	if (layout?.aggregation) out.aggregateByCategory = true;
	return out;
}

/** What an extended chart part is: its layout and, for the families the painter draws, a summary. */
export interface ChartExSummary {
	/** `cx:series/@layoutId` of the first series (`waterfall`, `clusteredColumn`, `funnel`...). */
	layoutId: string;
	summary?: ChartSummary;
}

const LEGEND_POSITIONS = ['r', 'l', 't', 'b', 'tr'];

/** The summary of a `cx:chartSpace`, or undefined when the part is not an extended chart. */
export function chartExSummary(source: string | XmlElement): ChartExSummary | undefined {
	const root =
		typeof source === 'string' ? parseXml(source, { label: 'Chart part' }).documentElement : source;
	if (!isChartExRoot(root)) return undefined;
	return chartExSummaryFromModel(parseChartExSpace(root).chartExSpace);
}

/** The text style Office draws an extended chart's title with (it has no `c:style`). */
function chartExFormatting(): NonNullable<ChartSummary['formatting']> {
	// Office draws it with the classic default text (grey, regular), not the larger bold text of
	// the newly authored style.
	return {
		builtInStyle: 2,
		sourceXml: '',
		entries: {
			title: {
				fontSize: 14,
				bold: false,
				textColor: {
					kind: 'scheme',
					value: 'tx1',
					transforms: [
						{ name: 'lumMod', value: '65000' },
						{ name: 'lumOff', value: '35000' },
					],
				},
				sourceXml: '',
			},
		},
	};
}

function summarySeries(
	space: ChartExSpace,
	source: ChartExSeries,
	chartType: import('./summary').ChartSummaryType,
): ChartSummarySeries {
	const data = seriesData(space, source);
	const categories = dimension(data, 'strDim', 'cat');
	const values = dimension(
		data,
		'numDim',
		chartType === 'treemap' || chartType === 'sunburst' ? 'size' : 'val',
	);
	const series: ChartSummarySeries = {
		categories: categories?.points ?? [],
		values: (values?.points ?? []).map((p) =>
			p !== '' && Number.isFinite(Number(p)) ? Number(p) : null,
		),
		...seriesFill(source),
	};
	if (chartType === 'funnel' || chartType === 'boxWhisker') {
		chartExSeriesPaint(source, series, chartType === 'funnel' && source.dataLabels !== undefined);
		if (source.hidden !== undefined) series.hidden = source.hidden;
	}
	if (chartType === 'boxWhisker') series.boxWhisker = source.layoutProperties ?? {};
	const name = source.tx?.value?.trim();
	if (name) series.name = name;
	const nameRef = source.tx?.formula?.formula.trim();
	if (nameRef) series.nameRef = nameRef;
	if (categories?.formula) series.categoriesRef = categories.formula;
	if (values?.formula) series.valuesRef = values.formula;
	return series;
}

/** The summary of a parsed extended chart part. */
export function chartExSummaryFromModel(space: ChartExSpace): ChartExSummary {
	const all = space.plotArea.series;
	const primary = all.find((s) => s.layoutId !== 'paretoLine') ?? all[0];
	const layoutId = primary?.layoutId ?? '';
	const chartType =
		layoutId === 'waterfall'
			? 'waterfall'
			: layoutId === 'clusteredColumn'
				? 'histogram'
				: layoutId === 'treemap' ||
					  layoutId === 'sunburst' ||
					  layoutId === 'funnel' ||
					  layoutId === 'boxWhisker'
					? layoutId
					: undefined;
	if (!primary || !chartType) return { layoutId };
	const series = summarySeries(space, primary, chartType);
	const layout = primary.layoutProperties;
	const legend = space.legend;
	const summary: ChartSummary = {
		chartType,
		series:
			chartType === 'boxWhisker'
				? all
						.filter((s) => s.layoutId === 'boxWhisker')
						.map((source) => summarySeries(space, source, chartType))
				: [series],
		showLegend: legend !== undefined,
	};
	if (chartType === 'boxWhisker') {
		summary.axes = space.plotArea.axes.flatMap((axis) => {
			const scaling = axis.scaling;
			if (!scaling) return [];
			const value = scaling.kind === 'val';
			return [
				{
					kind: value ? ('val' as const) : ('cat' as const),
					...(axis.id === undefined ? {} : { id: axis.id }),
					position: value ? 'l' : 'b',
					deleted: axis.hidden ?? false,
					scaling: {
						...(value && typeof scaling.min === 'number' ? { min: scaling.min } : {}),
						...(value && typeof scaling.max === 'number' ? { max: scaling.max } : {}),
					},
					majorGridlines: !!axis.majorGridlines,
					minorGridlines: !!axis.minorGridlines,
					...(axis.majorGridlines?.spPr ? { majorGridlinesSpPr: axis.majorGridlines.spPr } : {}),
					...(axis.minorGridlines?.spPr ? { minorGridlinesSpPr: axis.minorGridlines.spPr } : {}),
					...(value && typeof scaling.majorUnit === 'number'
						? { majorUnit: scaling.majorUnit }
						: {}),
					...(value && typeof scaling.minorUnit === 'number'
						? { minorUnit: scaling.minorUnit }
						: {}),
					...(axis.spPr ? { spPr: axis.spPr } : {}),
					...(axis.numberFormat ? { numberFormat: axis.numberFormat } : {}),
					...(axis.tickLabels ? {} : { tickLabelPosition: 'none' }),
				},
			];
		});
	}
	summary.formatting = chartExFormatting();
	const title = titleOf(space);
	if (title) summary.title = title;
	const position = legend?.position;
	if (position && LEGEND_POSITIONS.includes(position))
		summary.legendPosition = position as NonNullable<ChartSummary['legendPosition']>;
	if (chartType === 'waterfall') {
		summary.waterfall = {
			...(layout?.subtotals ? { subtotalIndices: [...layout.subtotals] } : {}),
			...(layout?.visibility?.connectorLines === false ? { connectorLines: false } : {}),
		};
	} else if (chartType === 'treemap' || chartType === 'sunburst') {
		const parentLayout = layout?.parentLabelLayout;
		summary.hierarchy = {
			levels: categoryLevels(seriesData(space, primary)),
			dataLabels: primary.dataLabels !== undefined,
			...(parentLayout === 'none' || parentLayout === 'banner' || parentLayout === 'overlapping'
				? { parentLabelLayout: parentLayout }
				: {}),
		};
	} else if (chartType === 'histogram') summary.histogram = histogramOptions(primary);
	return { layoutId, summary };
}
