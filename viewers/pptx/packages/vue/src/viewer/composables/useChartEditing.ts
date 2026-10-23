import {
	hideChartSeries,
	patchChartData as sharedPatchChartData,
	restoreFilteredSeries,
	setDataLabelsRangeCache,
} from 'ooxml-ui/pptx';
/**
 * useChartEditing: the framework-thin wiring layer for Vue's advanced chart
 * inspector. It owns the clone-mutate-emit plumbing so the `ChartPanel.vue`
 * subcomponents stay pure presentation.
 *
 * Every mutation funnels through one of two paths:
 *  - `replaceChartData` / `patchChartData` for plain shallow patches and the
 *    smart `chartDataChangeType` path, and
 *  - `useChartOpEditing`'s `applyChartOp` (a sibling file, split out to keep
 *    this one under the repo's 300-LOC guideline), which deep-clones the
 *    chart data and runs an in-place `pptx-viewer-core` SDK op against the
 *    clone so the live element is never mutated (Vue sees a fresh reference
 *    and history stays clean).
 *
 * The returned helpers all emit a SHALLOW `{ chartData }` patch via the
 * supplied `emitUpdate`, matching the inspector-panel `update` contract.
 */
import type {
	ChartPptxElement,
	PptxChartAxisFormatting,
	PptxChartData,
	PptxChartErrBars,
	PptxChartSeries,
	PptxChartStyle,
	PptxChartTrendline,
} from 'pptx-viewer-core';
import type { ComputedRef } from 'vue';

import { useChartOpEditing } from './useChartOpEditing';
import type { ChartOpEditing } from './useChartOpEditing';

export type {
	ChartAxisTitleStyleEdit,
	ChartGridlineStyleEdit,
	ChartMarkerEdit,
} from './useChartOpEditing';

/** The mutation helpers a chart inspector needs. */
export interface ChartEditing extends ChartOpEditing {
	patchChartData: (patch: Partial<PptxChartData>) => void;
	updateStyle: (patch: Partial<PptxChartStyle>) => void;
	updateAxis: (
		axisType: PptxChartAxisFormatting['axisType'],
		patch: Partial<PptxChartAxisFormatting>,
	) => void;
	setSeriesColor: (index: number, color: string | null) => void;
	setSeriesTrendline: (index: number, trendline: PptxChartTrendline | null) => void;
	setSeriesErrorBars: (index: number, errBars: PptxChartErrBars | null) => void;
	/** Patch a single series (e.g. rename) in place, preserving the rest. */
	updateSeries: (index: number, patch: Partial<PptxChartSeries>) => void;
	/** PowerPoint "Chart Filters": hide a currently-visible series. */
	hideSeries: (seriesIndex: number) => void;
	/** PowerPoint "Chart Filters": restore a series it hid. */
	restoreSeries: (filteredIndex: number) => void;
	/** "Value From Cells": edit one cached custom-label string for a series. */
	setLabelsRangeCache: (seriesIndex: number, pointIndex: number, text: string) => void;
}

/**
 * Build the chart-editing helper bundle.
 *
 * @param element    reactive accessor for the selected chart element (or null).
 * @param chartData  reactive accessor for that element's chart data (or null).
 * @param emitUpdate emits the shallow `{ chartData }` patch up to the host.
 */
export function useChartEditing(
	element: ComputedRef<ChartPptxElement | null>,
	chartData: ComputedRef<PptxChartData | null>,
	emitUpdate: (next: PptxChartData) => void,
): ChartEditing {
	const replaceChartData = (next: PptxChartData): void => emitUpdate(next);

	const patchChartData = (patch: Partial<PptxChartData>): void => {
		const data = chartData.value;
		if (!data) {
			return;
		}
		replaceChartData(sharedPatchChartData(data, patch));
	};

	const updateStyle = (patch: Partial<PptxChartStyle>): void => {
		const data = chartData.value;
		if (!data) {
			return;
		}
		replaceChartData({ ...data, style: { ...data.style, ...patch } });
	};

	const updateAxis = (
		axisType: PptxChartAxisFormatting['axisType'],
		patch: Partial<PptxChartAxisFormatting>,
	): void => {
		const data = chartData.value;
		if (!data) {
			return;
		}
		const axes = data.axes ? [...data.axes] : [];
		const index = axes.findIndex((a) => a.axisType === axisType);
		if (index === -1) {
			axes.push({ axisType, ...patch });
		} else {
			axes[index] = { ...axes[index], ...patch };
		}
		patchChartData({ axes });
	};

	const updateSeries = (index: number, patch: Partial<PptxChartSeries>): void => {
		const data = chartData.value;
		if (!data) {
			return;
		}
		const series = data.series.map((s, i) => (i === index ? { ...s, ...patch } : s));
		patchChartData({ series });
	};

	const setSeriesColor = (index: number, color: string | null): void =>
		updateSeries(index, { color: color ?? undefined });

	const setSeriesTrendline = (index: number, trendline: PptxChartTrendline | null): void =>
		updateSeries(index, { trendlines: trendline ? [trendline] : [] });

	// ── PowerPoint "Chart Filters" show/hide + "Value From Cells" cache edit ──
	const hideSeries = (seriesIndex: number): void => {
		const data = chartData.value;
		const next = data && hideChartSeries(data, seriesIndex);
		if (next) {
			replaceChartData(next);
		}
	};

	const restoreSeries = (filteredIndex: number): void => {
		const data = chartData.value;
		const next = data && restoreFilteredSeries(data, filteredIndex);
		if (next) {
			replaceChartData(next);
		}
	};

	const setLabelsRangeCache = (seriesIndex: number, pointIndex: number, text: string): void => {
		const data = chartData.value;
		if (!data) {
			return;
		}
		const seriesList = data.series.map((s, i) =>
			i === seriesIndex ? setDataLabelsRangeCache(s, pointIndex, text) : s,
		);
		patchChartData({ series: seriesList });
	};

	const setSeriesErrorBars = (index: number, errBars: PptxChartErrBars | null): void =>
		updateSeries(index, { errBars: errBars ? [errBars] : [] });

	const opEditing = useChartOpEditing(element, chartData, replaceChartData);

	return {
		patchChartData,
		updateStyle,
		updateAxis,
		setSeriesColor,
		setSeriesTrendline,
		setSeriesErrorBars,
		...opEditing,
		updateSeries,
		hideSeries,
		restoreSeries,
		setLabelsRangeCache,
	};
}
