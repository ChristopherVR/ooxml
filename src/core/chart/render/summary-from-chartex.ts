import { chartExSeriesPaint } from './summary-ex-series';
// Builds the painter's chart summary from an Office 2016 extended chart part (`cx:chartSpace`):
// the waterfall (`layoutId="waterfall"`), the histogram (`layoutId="clusteredColumn"`), the treemap
// and the sunburst. Word reads the cached points of `cx:data`; Excel files carry only the formula
// (a hidden `_xlchart.v1.N` defined name) and the host resolves it from live cells. Other extended
// families (funnel, box and whisker...) are reported by `layoutId` and not summarised. The part is
// read through the neutral chartex model (`chart/ex`).
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
				: layoutId === 'treemap' || layoutId === 'sunburst' || layoutId === 'funnel'
					? layoutId
					: undefined;
	if (!primary || !chartType) return { layoutId };
	const data = seriesData(space, primary);
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
		...seriesFill(primary),
	};
	if (chartType === 'funnel') chartExSeriesPaint(primary, series, true);
	const name = primary.tx?.value?.trim();
	if (name) series.name = name;
	const nameRef = primary.tx?.formula?.formula.trim();
	if (nameRef) series.nameRef = nameRef;
	if (categories?.formula) series.categoriesRef = categories.formula;
	if (values?.formula) series.valuesRef = values.formula;
	const layout = primary.layoutProperties;
	const legend = space.legend;
	const summary: ChartSummary = { chartType, series: [series], showLegend: legend !== undefined };
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
			levels: categoryLevels(data),
			dataLabels: primary.dataLabels !== undefined,
			...(parentLayout === 'none' || parentLayout === 'banner' || parentLayout === 'overlapping'
				? { parentLabelLayout: parentLayout }
				: {}),
		};
	} else if (chartType === 'histogram') summary.histogram = histogramOptions(primary);
	return { layoutId, summary };
}
