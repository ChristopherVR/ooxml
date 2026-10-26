import { setChartPointLabel, type ChartPointLabelEdit } from '../../../chart';
import type { PptxChartData, PptxChartDataLabel } from '../types';
import type { NeutralChartEdit } from './chart-neutral-save-plan';
import { same } from './chart-neutral-save-common';
import { neutralTextEdit } from './chart-neutral-save-text';
import { neutralShapeEdit } from './chart-neutral-save-shape';

const FIELDS = {
	deleted: 'deleted',
	showVal: 'showValue',
	showCatName: 'showCategoryName',
	showSerName: 'showSeriesName',
	showPercent: 'showPercent',
	showBubbleSize: 'showBubbleSize',
	showLegendKey: 'showLegendKey',
	position: 'position',
	separator: 'separator',
} as const;

function labelEdit(
	old: Partial<PptxChartDataLabel>,
	next: PptxChartDataLabel,
): ChartPointLabelEdit | undefined {
	const remaining = structuredClone(next);
	const edit: ChartPointLabelEdit = {};
	for (const [source, target] of Object.entries(FIELDS) as [
		keyof typeof FIELDS,
		(typeof FIELDS)[keyof typeof FIELDS],
	][]) {
		if (same(old[source], next[source])) continue;
		if (next[source] === undefined) return undefined;
		Object.assign(edit, { [target]: next[source] });
		if (old[source] === undefined) delete remaining[source];
		else Object.assign(remaining, { [source]: old[source] });
	}
	if (old.text !== next.text || old.textFromCells !== next.textFromCells) {
		edit.text = next.textFromCells ? null : (next.text ?? null);
		remaining.text = old.text;
		remaining.textFromCells = old.textFromCells;
	}
	if (!same(old.txPr, next.txPr)) {
		edit.font = neutralTextEdit(old.txPr, next.txPr);
		remaining.txPr = old.txPr;
	}
	if (!same(old.spPr, next.spPr)) {
		edit.shape = neutralShapeEdit(old.spPr, next.spPr);
		remaining.spPr = old.spPr;
	}
	if (old.numberFormat !== next.numberFormat) {
		if (next.numberFormat === undefined) return undefined;
		edit.numberFormat = { formatCode: next.numberFormat, sourceLinked: false };
		remaining.numberFormat = old.numberFormat;
	}
	if (!same(old.layout, next.layout)) {
		if (!same(old.layout?.ext, next.layout?.ext)) return undefined;
		const { ext: _ext, ...layout } = next.layout ?? {};
		edit.layout = next.layout ? layout : null;
		remaining.layout = old.layout;
	}
	return same({ ...remaining, idx: old.idx }, old) ? edit : undefined;
}

export function planNeutralPointLabels(
	before: PptxChartData,
	probe: PptxChartData,
	edits: NeutralChartEdit[],
): void {
	for (const [seriesIndex, series] of probe.series.entries()) {
		const old = before.series[seriesIndex];
		if (!old || same(old.dataLabels, series.dataLabels)) continue;
		const pending: NeutralChartEdit[] = [];
		let supported = true;
		for (const label of series.dataLabels ?? []) {
			const previous = old.dataLabels?.find((item) => item.idx === label.idx) ?? {};
			const edit = labelEdit(previous, label);
			if (!edit) {
				supported = false;
				break;
			}
			pending.push((space) => setChartPointLabel(space, seriesIndex, label.idx, edit));
		}
		if (!supported) continue;
		for (const previous of old.dataLabels ?? [])
			if (!series.dataLabels?.some((item) => item.idx === previous.idx))
				pending.push((space) => setChartPointLabel(space, seriesIndex, previous.idx, null));
		edits.push(...pending);
		if (old.dataLabels === undefined) delete series.dataLabels;
		else series.dataLabels = old.dataLabels;
	}
}
