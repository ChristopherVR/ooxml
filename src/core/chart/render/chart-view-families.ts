// The view-model parts of the bubble and hierarchy families: bubble sizes resolved from live or
// cached values, and the category levels of a treemap or sunburst. The geometry is the shared code
// PowerPoint also uses (`../bubble-radius`, `../hierarchy`).
import { hierarchyBranchLabels, normalizeLevels } from '../hierarchy';
import type { ChartViewModel } from './chart-view';
import { resolveRef, toNumber, toText, type EvaluateRef } from './chart-view-series';
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

const columnNumber = (letters: string): number =>
	[...letters].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);

/** The number of columns a range reference such as `Data!$T$2:$U$6` spans (1 when unknown). */
export function refColumns(ref: string): number {
	const match = /\$?([A-Z]{1,3})\$?\d+:\$?([A-Z]{1,3})\$?\d+\s*$/i.exec(ref);
	if (!match) return 1;
	return Math.max(
		1,
		columnNumber((match[2] ?? '').toUpperCase()) - columnNumber((match[1] ?? '').toUpperCase()) + 1,
	);
}

/**
 * A spreadsheet hierarchy lives in one multi-column range (outermost group first, leaf last).
 * Resolved live it comes back row-major; this splits it into levels, leaf first.
 */
function liveLevels(ref: string | undefined, evaluateRef: EvaluateRef): string[][] | undefined {
	const columns = ref ? refColumns(ref) : 1;
	const flat = columns > 1 ? resolveRef(ref, evaluateRef)?.map(toText) : undefined;
	if (!flat || flat.length % columns !== 0) return undefined;
	const levels: string[][] = Array.from({ length: columns }, () => []);
	flat.forEach((label, i) => levels[columns - 1 - (i % columns)]?.push(label));
	return levels;
}

/** Treemap and sunburst: the category levels (parents filled down) and the legend's branches. */
export function hierarchyView<S extends ChartSummarySeries>(
	chart: ChartSummary<S>,
	model: ChartViewModel,
	evaluateRef: EvaluateRef,
): void {
	const first = model.series[0];
	const live = liveLevels(chart.series[0]?.categoriesRef, evaluateRef);
	if (live?.[0]) model.categories = live[0];
	const count = Math.max(first?.values.length ?? 0, model.categories.length);
	const supplied = live ?? chart.hierarchy?.levels ?? [];
	const levels = supplied.length > 1 ? normalizeLevels(supplied, count) : [];
	model.hierarchy = {
		levels,
		parentLabelLayout: chart.hierarchy?.parentLabelLayout ?? 'banner',
		labels: chart.hierarchy?.dataLabels !== false,
		branches: hierarchyBranchLabels(
			model.series.map((s) => ({ name: s.name, values: s.values.map((v) => v ?? 0) })),
			model.categories,
			levels.length > 1 ? levels : undefined,
		),
	};
}
