/**
 * The chart elements a PowerPoint Format pane can target, in the order
 * PowerPoint lists them in Chart Format > Current Selection and in the pane's
 * element picker. Captions and pane titles come from the shared chart Format
 * pane (`src/ui/src/chart/format-pane-targets.ts`), whose `chartFormat.*` keys
 * this module maps onto PowerPoint's `pptx.chartFormat.*` dictionary.
 *
 * @module render/chart-format/chart-format-targets
 */
import type {
	PptxChartAxisFormatting,
	PptxChartData,
	PptxChartFormatTarget,
	PptxChartFormatTargetKey,
} from 'ooxml-core/pptx';
import { parseChartFormatTargetKey } from 'ooxml-core/pptx';

import { translateChartFormatText as translateNeutral } from '../../../chart/format-pane-common';
import {
	CHART_AREA_TARGET,
	chartFormatPaneTitle as neutralPaneTitle,
	chartFormatTargetLabel as neutralTargetLabel,
	chartFormatTargetOptions,
	type ChartFormatDescribe,
} from '../../../chart/format-pane-targets';
import type { ChartFormatTargetOption, ChartFormatText } from '../../../chart/format-pane-types';

export type { ChartFormatTargetOption, ChartFormatText } from '../../../chart/format-pane-types';

/** The PowerPoint dictionary key of a shared chart Format caption key. */
export const pptxChartFormatKey = (key: string): string =>
	key.startsWith('chartFormat.') ? `pptx.${key}` : key;

/** Whether an axis is drawn vertically (left/right edge). */
function isVerticalAxis(axis: PptxChartAxisFormatting, data: PptxChartData): boolean {
	if (axis.axPos) {
		return axis.axPos === 'l' || axis.axPos === 'r';
	}
	const horizontalBars = data.barDirection === 'bar';
	return axis.axisType === 'valAx' ? !horizontalBars : horizontalBars;
}

const AXIS_KINDS = { catAx: 'cat', valAx: 'val', dateAx: 'date', serAx: 'ser' } as const;

/** The facts captions need about `data`. */
export function describePptxChart(data: PptxChartData): ChartFormatDescribe {
	return {
		axis: (index) => {
			const axis = data.axes?.[index];
			if (!axis) return undefined;
			return {
				kind: AXIS_KINDS[axis.axisType] ?? 'cat',
				vertical: isVerticalAxis(axis, data),
				secondary: (data.axes ?? []).slice(0, index).some((a) => a.axisType === axis.axisType),
			};
		},
		seriesName: (index) => data.series[index]?.name || `Series ${index + 1}`,
	};
}

/** The picker caption of `target`. */
export function chartFormatTargetLabel(
	data: PptxChartData,
	target: PptxChartFormatTarget,
): ChartFormatText {
	return neutralTargetLabel(target, describePptxChart(data));
}

/** The pane heading for `target` ("Format Chart Area", "Format Axis", ...). */
export const chartFormatPaneTitle = (target: PptxChartFormatTarget): ChartFormatText =>
	neutralPaneTitle(target);

function hasTitle(data: PptxChartData): boolean {
	if (data.style?.hasTitle === false) {
		return false;
	}
	return Boolean(data.style?.hasTitle || data.title || data.titleRuns?.length);
}

/** Every element of `data` a Format pane can target, in PowerPoint's picker order. */
export function listChartFormatTargets(data: PptxChartData): ChartFormatTargetOption[] {
	const targets: PptxChartFormatTarget[] = [{ kind: 'chartArea' }];
	if (hasTitle(data)) {
		targets.push({ kind: 'title' });
	}
	(data.axes ?? []).forEach((axis, axisIndex) => {
		if (axis.deleted) {
			return;
		}
		targets.push({ kind: 'axis', axisIndex });
		if (axis.majorGridlines) {
			targets.push({ kind: 'majorGridlines', axisIndex });
		}
		if (axis.minorGridlines) {
			targets.push({ kind: 'minorGridlines', axisIndex });
		}
	});
	if (data.style?.hasLegend) {
		targets.push({ kind: 'legend' });
	}
	targets.push({ kind: 'plotArea' });
	if (data.style?.hasDataLabels) {
		targets.push({ kind: 'dataLabels' });
	}
	data.series.forEach((_, seriesIndex) => {
		targets.push({ kind: 'series', seriesIndex });
		// Each series' own labels (PowerPoint's "Series "X" Data Labels"), listed even while
		// hidden: their Label Options are where one series' labels are switched on.
		targets.push({ kind: 'dataLabels', seriesIndex });
	});
	return chartFormatTargetOptions(targets, describePptxChart(data));
}

/** A valid target of `data` the picker does not list (a deleted axis), or undefined. */
export function resolveUnlistedChartFormatTarget(
	data: PptxChartData,
	key: PptxChartFormatTargetKey,
): ChartFormatTargetOption | undefined {
	const parsed = parseChartFormatTargetKey(key);
	if (!parsed) return undefined;
	const valid =
		(parsed.kind === 'series' && data.series[parsed.seriesIndex]) ||
		(parsed.kind === 'dataLabels' &&
			parsed.seriesIndex !== undefined &&
			data.series[parsed.seriesIndex]) ||
		((parsed.kind === 'axis' ||
			parsed.kind === 'majorGridlines' ||
			parsed.kind === 'minorGridlines') &&
			data.axes?.[parsed.axisIndex]);
	return valid ? { key, target: parsed, label: chartFormatTargetLabel(data, parsed) } : undefined;
}

/**
 * `key` when it names an element `data` still has, else the chart area (the
 * target a stale key falls back to after the element was removed).
 */
export function resolveChartFormatTarget(
	data: PptxChartData | undefined,
	key: PptxChartFormatTargetKey | undefined,
): ChartFormatTargetOption {
	if (!data || !key) return CHART_AREA_TARGET;
	return (
		listChartFormatTargets(data).find((option) => option.key === key) ??
		resolveUnlistedChartFormatTarget(data, key) ??
		CHART_AREA_TARGET
	);
}

/** Resolve a {@link ChartFormatText} with a PowerPoint translator that echoes unknown keys. */
export function translateChartFormatText(
	value: ChartFormatText,
	t: (key: string, params?: Readonly<Record<string, string | number>>) => string,
): string {
	return translateNeutral(value, (key, params) => {
		const pptxKey = pptxChartFormatKey(key);
		const translated = t(pptxKey, params);
		return translated === pptxKey ? key : translated;
	});
}
