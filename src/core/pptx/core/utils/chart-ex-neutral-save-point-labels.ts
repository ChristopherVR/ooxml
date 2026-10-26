import { editChartExPointLabel, type ChartExPointLabelEdit } from '../../../chart';
import type { PptxChartData, PptxChartDataLabel } from '../types';
import type { ChartExEdit } from './chart-ex-neutral-save';
import { same } from './chart-neutral-save-common';
import { neutralTextEdit } from './chart-neutral-save-text';
import { neutralShapeEdit } from './chart-neutral-save-shape';
const flags = { showVal: 'value', showCatName: 'categoryName', showSerName: 'seriesName' } as const;
function labelEdit(
	old: Partial<PptxChartDataLabel>,
	next: PptxChartDataLabel,
): ChartExPointLabelEdit | undefined {
	const remaining = structuredClone(next),
		edit: ChartExPointLabelEdit = {};
	for (const [key, target] of Object.entries(flags) as [
		keyof typeof flags,
		(typeof flags)[keyof typeof flags],
	][])
		if (old[key] !== next[key]) {
			edit.visibility ??= {};
			edit.visibility[target] = next[key] ?? false;
			remaining[key] = old[key];
		}
	for (const key of ['deleted', 'position', 'separator'] as const)
		if (old[key] !== next[key]) {
			if (next[key] === undefined) return undefined;
			Object.assign(edit, { [key]: next[key] });
			Object.assign(remaining, { [key]: old[key] });
		}
	if (old.numberFormat !== next.numberFormat) {
		if (next.numberFormat === undefined) return undefined;
		edit.numberFormat = { formatCode: next.numberFormat, sourceLinked: false };
		remaining.numberFormat = old.numberFormat;
	}
	if (!same(old.txPr, next.txPr)) {
		edit.font = neutralTextEdit(old.txPr, next.txPr);
		remaining.txPr = old.txPr;
	}
	if (!same(old.spPr, next.spPr)) {
		edit.shape = neutralShapeEdit(old.spPr, next.spPr);
		remaining.spPr = old.spPr;
	}
	return same({ ...remaining, idx: old.idx }, old) ? edit : undefined;
}
export function planChartExPointLabels(
	before: PptxChartData,
	probe: PptxChartData,
	edits: ChartExEdit[],
): void {
	for (const [seriesIndex, series] of probe.series.entries()) {
		const old = before.series[seriesIndex];
		if (!old || same(old.dataLabels, series.dataLabels)) continue;
		const pending: ChartExEdit[] = [];
		let supported = true;
		for (const label of series.dataLabels ?? []) {
			const previous = old.dataLabels?.find((item) => item.idx === label.idx) ?? {};
			const edit = labelEdit(previous, label);
			if (!edit) {
				supported = false;
				break;
			}
			pending.push((space) => editChartExPointLabel(space, seriesIndex, label.idx, edit));
		}
		if (!supported) {
			// A custom-text or manual-layout edit has no ChartEx node; do not silently ignore it.
			edits.push(() => {
				throw new RangeError('The ChartEx point-label edit is not supported.');
			});
		} else {
			for (const label of old.dataLabels ?? [])
				if (!series.dataLabels?.some((item) => item.idx === label.idx))
					pending.push((space) => editChartExPointLabel(space, seriesIndex, label.idx, null));
			edits.push(...pending);
		}
		if (old.dataLabels === undefined) delete series.dataLabels;
		else series.dataLabels = old.dataLabels;
	}
}
