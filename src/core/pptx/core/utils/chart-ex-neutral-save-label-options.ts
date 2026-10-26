import { editChartExLabels, type ChartExLabelsEdit } from '../../../chart';
import type { PptxChartData, PptxChartDataLabelOptions } from '../types';
import type { ChartExEdit } from './chart-ex-neutral-save';
import { same, revertStyle } from './chart-neutral-save-common';
import { neutralTextEdit } from './chart-neutral-save-text';
import { neutralShapeEdit } from './chart-neutral-save-shape';
function options(
	before: PptxChartDataLabelOptions = {},
	after: PptxChartDataLabelOptions = {},
): ChartExLabelsEdit {
	const edit: ChartExLabelsEdit = {};
	for (const [key, target] of [
		['showValue', 'value'],
		['showCategory', 'categoryName'],
		['showSeriesName', 'seriesName'],
	] as const)
		if (before[key] !== after[key]) {
			edit.visibility ??= {};
			edit.visibility[target] = after[key] ?? null;
		}
	for (const key of ['position', 'separator'] as const)
		if (before[key] !== after[key]) edit[key] = after[key] ?? null;
	if (before.numberFormat !== after.numberFormat)
		edit.numberFormat =
			after.numberFormat === undefined
				? null
				: { formatCode: after.numberFormat, sourceLinked: false };
	if (!same(before.txPr, after.txPr)) edit.font = neutralTextEdit(before.txPr, after.txPr);
	if (!same(before.labelShape, after.labelShape))
		edit.shape = after.labelShape ? neutralShapeEdit(before.labelShape, after.labelShape) : null;
	const supported = new Set([
		'showValue',
		'showCategory',
		'showSeriesName',
		'position',
		'separator',
		'numberFormat',
		'txPr',
		'labelShape',
	]);
	for (const key of new Set([...Object.keys(before), ...Object.keys(after)]))
		if (
			!supported.has(key) &&
			!same(before[key as keyof typeof before], after[key as keyof typeof after])
		) {
			if (
				[
					'showPercent',
					'showBubbleSize',
					'showLegendKey',
					'showLeaderLines',
					'extLeaderLines',
				].includes(key) &&
				after[key as keyof typeof after] !== true
			)
				continue;
			throw new RangeError(`ChartEx does not support the ${key} label option.`);
		}
	return edit;
}
export function planChartExLabelOptions(
	before: PptxChartData,
	probe: PptxChartData,
	edits: ChartExEdit[],
): void {
	const plan = (
		old: PptxChartDataLabelOptions | undefined,
		next: PptxChartDataLabelOptions | undefined,
		index?: number,
	) => {
		if (same(old, next)) return;
		// Build inside the operation so an unsupported setting becomes a typed host refusal.
		edits.push((space) => editChartExLabels(space, index, options(old, next)));
	};
	if (!same(before.style?.dataLabels, probe.style?.dataLabels)) {
		plan(before.style?.dataLabels, probe.style?.dataLabels);
		revertStyle(probe, before, ['dataLabels']);
	}
	probe.series.forEach((series, index) => {
		const old = before.series[index];
		if (!old || same(old.dataLabelOptions, series.dataLabelOptions)) return;
		plan(old.dataLabelOptions, series.dataLabelOptions, index);
		if (old.dataLabelOptions === undefined) delete series.dataLabelOptions;
		else series.dataLabelOptions = old.dataLabelOptions;
	});
}
