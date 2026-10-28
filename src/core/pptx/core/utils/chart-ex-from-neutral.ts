// ChartEx facts come from the namespace-aware shared parser. Host-resolved paint and fonts are
// supplied separately so PowerPoint retains its theme and raw geography-cache compatibility.
import type {
	ChartExData,
	ChartExDataLabel,
	ChartExDimension,
	ChartExLabelVisibility,
	ChartExSpace,
	ChartExText,
} from '../../../chart';
import type {
	PptxChartAxisFormatting,
	PptxChartData,
	PptxChartDataLabelOptions,
	PptxChartDataLabelPosition,
	PptxChartSeries,
} from '../types';
import { chartExFamilyFields } from './chart-ex-neutral-families';

/** The first authored layout determines PowerPoint's ChartEx family. */
export function chartExTypeFromNeutral(space: ChartExSpace): PptxChartData['chartType'] {
	const types: Record<string, PptxChartData['chartType']> = {
		waterfall: 'waterfall',
		funnel: 'funnel',
		treemap: 'treemap',
		sunburst: 'sunburst',
		boxWhisker: 'boxWhisker',
		clusteredColumn: 'histogram',
		paretoLine: 'histogram',
		regionMap: 'regionMap',
		histogram: 'histogram',
	};
	return types[space.plotArea.series[0]?.layoutId ?? ''] ?? 'funnel';
}

type Styled = {
	categories: string[];
	series: PptxChartSeries[];
	chartData?: Partial<PptxChartData>;
};
const positions = new Set(['bestFit', 'b', 'ctr', 'inBase', 'inEnd', 'l', 'outEnd', 'r', 't']);
const position = (value: string | undefined) =>
	positions.has(value ?? '') ? (value as PptxChartDataLabelPosition) : undefined;
const units: Record<string, number> = {
	hundreds: 100,
	thousands: 1000,
	tenThousands: 10000,
	hundredThousands: 100000,
	millions: 1e6,
	tenMillions: 1e7,
	hundredMillions: 1e8,
	billions: 1e9,
	trillions: 1e12,
};
export function chartExText(tx: ChartExText | undefined): string | undefined {
	const runs = tx?.rich?.paragraphs.flatMap((p) =>
		p.runs.filter((run) => run.text !== '\n').map((run) => run.text),
	);
	return runs?.length ? runs.join('') : tx?.value?.trim() || undefined;
}
function dimension(data: ChartExData | undefined, kind: ChartExDimension['kind'], type: string) {
	const map = new Map<string, ChartExDimension>();
	for (const d of data?.dimensions ?? [])
		if (d.kind === kind) map.set(d.type?.trim() || (kind === 'numDim' ? 'val' : 'cat'), d);
	return map.get(type) ?? map.values().next().value;
}
const levelText = (dimension: ChartExDimension | undefined) =>
	dimension?.levels.map((level) => level.points.map((point) => point.value.trim()));
function flags(visibility: ChartExLabelVisibility | undefined) {
	return {
		showVal: visibility ? visibility.value === true : undefined,
		showCatName: visibility ? visibility.categoryName === true : undefined,
		showSerName: visibility ? visibility.seriesName === true : undefined,
	};
}
/** Adapt shared facts, retaining each series' host-resolved styles. No package state is read here. */
export function chartExFromNeutral(space: ChartExSpace, styled: Styled) {
	const dataOf = (index: number) => {
		const series = space.plotArea.series[index]!;
		return series.data ?? space.chartData.data.find((data) => data.id === series.dataId);
	};
	const levels = levelText(
		dimension(space.plotArea.series.length ? dataOf(0) : undefined, 'strDim', 'cat'),
	);
	const series = space.plotArea.series.map((source, index): PptxChartSeries => {
		const result = { ...styled.series[index] };
		const values = levelText(dimension(dataOf(index), 'numDim', 'val'))?.[0]
			?.map((value) => Number.parseFloat(value))
			.filter(Number.isFinite);
		result.name = source.tx?.value?.trim() || `Series ${index + 1}`;
		result.values = values?.length ? values : [0];
		for (const key of [
			'histogramOptions',
			'waterfallOptions',
			'boxWhiskerOptions',
			'treemapOptions',
		] as const)
			delete result[key];
		Object.assign(result, chartExFamilyFields(source));
		const group = source.dataLabels;
		if (!group) {
			delete result.dataLabels;
			delete result.dataLabelOptions;
		} else {
			const options: PptxChartDataLabelOptions = { ...result.dataLabelOptions };
			for (const [key, target] of [
				['value', 'showValue'],
				['categoryName', 'showCategory'],
				['seriesName', 'showSeriesName'],
			] as const)
				if (group.visibility?.[key] !== undefined) options[target] = group.visibility[key];
			delete options.position;
			delete options.numberFormat;
			const pos = position(group.position),
				format = group.numberFormat?.formatCode.trim();
			if (pos) options.position = pos;
			if (format) options.numberFormat = format;
			if (Object.keys(options).length) result.dataLabelOptions = options;
			else delete result.dataLabelOptions;
			if (group.labels.length || group.hidden.length)
				result.dataLabels = [
					...group.labels,
					...group.hidden
						.filter((index) => !group.labels.some((label) => label.index === index))
						.map((index): ChartExDataLabel => ({ index })),
				].map((label) => {
					const styled = result.dataLabels?.find((item) => item.idx === label.index);
					const out = {
						...styled,
						idx: label.index ?? 0,
						...flags(label.visibility ?? group.visibility),
					};
					if (group.hidden.includes(label.index ?? 0)) out.deleted = true;
					delete out.position;
					delete out.numberFormat;
					const pos = position(label.position),
						format = label.numberFormat?.formatCode.trim();
					if (pos) out.position = pos;
					if (format) out.numberFormat = format;
					return out;
				});
			else delete result.dataLabels;
		}
		return result;
	});
	const axes = space.plotArea.axes.map((source, index): PptxChartAxisFormatting => {
		const out = {
			...styled.chartData?.axes?.[index],
			axisType: source.scaling?.kind === 'cat' ? 'catAx' : 'valAx',
		} as PptxChartAxisFormatting;
		for (const key of [
			'axisId',
			'deleted',
			'titleText',
			'numFmt',
			'majorGridlines',
			'minorGridlines',
			'tickLblPos',
			'displayUnits',
			'displayUnitsValue',
			'displayUnitsLabel',
		] as const)
			delete out[key];
		if (source.id !== undefined) out.axisId = source.id;
		for (const key of ['min', 'max', 'majorUnit', 'minorUnit'] as const) {
			delete out[key];
			if (source.scaling?.kind === 'val' && typeof source.scaling[key] === 'number')
				out[key] = source.scaling[key];
		}
		if (source.hidden) out.deleted = true;
		const title = chartExText(source.title?.tx);
		if (title) out.titleText = title;
		const format = source.numberFormat?.formatCode.trim();
		if (format)
			out.numFmt = { formatCode: format, sourceLinked: source.numberFormat?.sourceLinked === true };
		if (source.majorGridlines) out.majorGridlines = true;
		if (source.minorGridlines) out.minorGridlines = true;
		if (!source.tickLabels) out.tickLblPos = 'none';
		const unit = source.units?.unit;
		const divisor = unit ? (units[unit] ?? Number.parseFloat(unit)) : NaN;
		if (Number.isFinite(divisor) && divisor > 0) {
			out.displayUnits = 'custom';
			out.displayUnitsValue = divisor;
			const old = styled.chartData?.axes?.[index]?.displayUnitsLabel;
			const label = typeof old === 'object' ? { ...old } : {};
			const text = chartExText(source.units?.label?.tx);
			if (text) label.text = text;
			if (Object.keys(label).length) out.displayUnitsLabel = label;
		}
		return out;
	});
	const title = chartExText(space.title?.tx);
	const chartData = { ...styled.chartData };
	if (space.plotArea.series[0]?.layoutId === 'funnel') {
		const category = space.plotArea.axes.find((axis) => axis.scaling?.kind === 'cat')?.scaling;
		if (category?.kind === 'cat' && typeof category.gapWidth === 'number')
			chartData.funnelGapWidth = category.gapWidth;
		else delete chartData.funnelGapWidth;
	}
	if (axes.length) chartData.axes = axes;
	else delete chartData.axes;
	if (title) chartData.title = title;
	else delete chartData.title;
	return {
		categories: (levels?.[0] ?? []).filter(Boolean),
		categoryLevels: levels,
		series,
		hasDataLabels: space.plotArea.series.some((series) => {
			const v = series.dataLabels?.visibility;
			return v?.value === true || v?.categoryName === true || v?.seriesName === true;
		}),
		...(Object.keys(chartData).length ? { chartData } : {}),
	};
}
