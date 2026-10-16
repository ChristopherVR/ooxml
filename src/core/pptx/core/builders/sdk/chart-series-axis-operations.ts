/**
 * PowerPoint's Format Data Series > Series Options > Plot Series On
 * (Primary Axis / Secondary Axis).
 *
 * In ChartML a series plots against the value axis its chart group names in
 * `c:axId`. Moving a series to the secondary axis therefore needs a secondary
 * value axis (on the far side, crossing at the maximum) and a hidden
 * secondary category axis paired with it; this creates both when the chart
 * has none, points the series' `axisId` at the new value axis, and the save
 * path gives the series a chart group of its own. Moving the last series back
 * removes the pair again, as PowerPoint does.
 *
 * @module sdk/chart-series-axis-operations
 */
import type {
	PptxChartAxisFormatting,
	PptxChartData,
	PptxChartElementFormats,
	PptxChartType,
} from '../../types';

/** Chart types PowerPoint offers a secondary axis for (2-D cartesian groups). */
const SECONDARY_TYPES = new Set<PptxChartType>([
	'bar',
	'line',
	'area',
	'scatter',
	'bubble',
	'combo',
]);

/** Whether `data` can plot a series on a secondary axis. */
export function chartSupportsSecondaryAxis(data: PptxChartData): boolean {
	return (
		SECONDARY_TYPES.has(data.chartType) &&
		(data.axes ?? []).some((axis) => axis.axisType === 'valAx')
	);
}

function valueAxes(data: PptxChartData): PptxChartAxisFormatting[] {
	return (data.axes ?? []).filter((axis) => axis.axisType === 'valAx');
}

/**
 * The primary and secondary value axes. A scatter chart's two primary axes
 * are both value axes, so a value axis counts as secondary only when it is on
 * the far side (`r` / `t`) and another value axis shares its orientation.
 */
function valueAxisPair(data: PptxChartData): {
	primary?: PptxChartAxisFormatting;
	secondary?: PptxChartAxisFormatting;
} {
	const values = valueAxes(data);
	const vertical = (axis: PptxChartAxisFormatting) => axis.axPos === 'l' || axis.axPos === 'r';
	const primary =
		values.find((axis) => axis.axPos === 'l') ??
		values.find((axis) => axis.axPos === 'b' && data.barDirection === 'bar') ??
		values[0];
	const secondary = values.find(
		(axis) =>
			axis !== primary &&
			(axis.axPos === 'r' || axis.axPos === 't') &&
			primary !== undefined &&
			vertical(axis) === vertical(primary),
	);
	return { primary, secondary };
}

/** Whether series `seriesIndex` is plotted against the secondary value axis. */
export function isChartSeriesOnSecondaryAxis(data: PptxChartData, seriesIndex: number): boolean {
	const series = data.series[seriesIndex];
	const { secondary } = valueAxisPair(data);
	return Boolean(series && secondary && series.axisId === secondary.axisId);
}

/** Rename `axis:N` style keys after removing the axes at `removed` indices. */
function reindexAxisFormats(
	formats: PptxChartElementFormats | undefined,
	removed: readonly number[],
): PptxChartElementFormats | undefined {
	if (!formats) return formats;
	const out: PptxChartElementFormats = {};
	for (const [key, value] of Object.entries(formats)) {
		const match = /^(axis|majorGridlines|minorGridlines):(\d+)$/u.exec(key);
		if (!match) {
			out[key] = value;
			continue;
		}
		const index = Number(match[2]);
		if (removed.includes(index)) continue;
		const shift = removed.filter((r) => r < index).length;
		out[`${match[1]}:${index - shift}`] = value;
	}
	return out;
}

/**
 * Plot series `seriesIndex` on the secondary (`true`) or primary (`false`)
 * value axis. Returns `data` unchanged when the chart cannot have a
 * secondary axis or the series is already there.
 *
 * @example
 * ```ts
 * const next = withChartSeriesOnSecondaryAxis(chart.chartData, 1, true);
 * ```
 */
export function withChartSeriesOnSecondaryAxis(
	data: PptxChartData,
	seriesIndex: number,
	secondary: boolean,
): PptxChartData {
	const series = data.series[seriesIndex];
	if (!series || !chartSupportsSecondaryAxis(data)) return data;
	if (isChartSeriesOnSecondaryAxis(data, seriesIndex) === secondary) return data;
	const axes = [...(data.axes ?? [])];
	const pair = valueAxisPair(data);
	const primary = pair.primary;
	if (!primary || primary.axisId === undefined) return data;
	if (!secondary) {
		const nextSeries = data.series.map((s, i) =>
			i === seriesIndex ? { ...s, axisId: primary.axisId } : s,
		);
		const stillUsed = nextSeries.some((s) => s.axisId === pair.secondary?.axisId);
		if (stillUsed || !pair.secondary) {
			return { ...data, series: nextSeries };
		}
		// The last series left: drop the secondary value axis and its hidden partner.
		const drop = new Set([pair.secondary.axisId, pair.secondary.crossAxisId]);
		const removed = axes.flatMap((axis, i) => (drop.has(axis.axisId) ? [i] : []));
		return {
			...data,
			series: nextSeries,
			axes: axes.filter((_, i) => !removed.includes(i)),
			elementFormats: reindexAxisFormats(data.elementFormats, removed),
		};
	}
	let target = pair.secondary;
	if (!target) {
		const used = axes.map((axis) => axis.axisId ?? 0);
		const valueId = Math.max(0, ...used) + 1;
		const crossId = valueId + 1;
		const primaryCross = axes.find((axis) => axis.axisId === primary.crossAxisId);
		const horizontal = primary.axPos === 'b' || primary.axPos === 't';
		target = {
			axisType: 'valAx',
			axisId: valueId,
			crossAxisId: crossId,
			axPos: horizontal ? 't' : 'r',
			crosses: 'max',
			crossBetween: primary.crossBetween ?? 'between',
			majorTickMark: 'out',
			minorTickMark: 'none',
			tickLblPos: 'nextTo',
			...(primary.numFmt ? { numFmt: { ...primary.numFmt } } : {}),
		};
		const cross: PptxChartAxisFormatting = {
			axisType: primaryCross?.axisType ?? 'catAx',
			axisId: crossId,
			crossAxisId: valueId,
			axPos: primaryCross?.axPos ?? (horizontal ? 'l' : 'b'),
			deleted: true,
			crosses: 'autoZero',
			majorTickMark: 'out',
			minorTickMark: 'none',
			tickLblPos: 'nextTo',
		};
		axes.push(target, cross);
	}
	return {
		...data,
		axes,
		series: data.series.map((s, i) => (i === seriesIndex ? { ...s, axisId: target!.axisId } : s)),
	};
}
