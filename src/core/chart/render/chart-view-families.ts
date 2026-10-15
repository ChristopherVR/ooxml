// The view-model parts of the bubble and hierarchy families: bubble sizes resolved from live or
// cached values, and the category levels of a treemap or sunburst. The geometry is the shared code
// PowerPoint also uses (`../bubble-radius`, `../hierarchy`).
import { hierarchyBranchLabels, normalizeLevels } from '../hierarchy';
import type { ChartViewModel } from './chart-view';
import { resolveRef, toNumber, type EvaluateRef } from './chart-view-series';
import type { ChartSummary, ChartSummarySeries } from './summary';

/** Bubble sizes per series, from `c:bubbleSize` references when the host resolves them. */
export function bubbleView<S extends ChartSummarySeries>(
	chart: ChartSummary<S>,
	model: ChartViewModel,
	evaluateRef: EvaluateRef,
): void {
	chart.series.forEach((source, i) => {
		const view = model.series[i];
		if (!view) return;
		const live = resolveRef(source.bubbleSizesRef, evaluateRef)?.map(toNumber);
		const sizes = live ?? source.bubbleSizes;
		if (sizes) view.sizes = sizes;
	});
	model.bubble = {
		sizeRepresents: chart.bubble?.sizeRepresents ?? 'area',
		...(chart.bubble?.bubbleScale === undefined ? {} : { bubbleScale: chart.bubble.bubbleScale }),
		showNegativeBubbles: chart.bubble?.showNegativeBubbles === true,
	};
}

/** Treemap and sunburst: the category levels (parents filled down) and the legend's branches. */
export function hierarchyView<S extends ChartSummarySeries>(
	chart: ChartSummary<S>,
	model: ChartViewModel,
): void {
	const first = model.series[0];
	const count = Math.max(first?.values.length ?? 0, model.categories.length);
	const supplied = chart.hierarchy?.levels ?? [];
	const levels = supplied.length > 1 ? normalizeLevels(supplied, count) : [];
	model.hierarchy = {
		levels,
		parentLabelLayout: chart.hierarchy?.parentLabelLayout ?? 'banner',
		branches: hierarchyBranchLabels(
			model.series.map((s) => ({ name: s.name, values: s.values.map((v) => v ?? 0) })),
			model.categories,
			levels.length > 1 ? levels : undefined,
		),
	};
}
