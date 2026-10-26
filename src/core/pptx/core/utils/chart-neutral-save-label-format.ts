import { setChartDataLabelFormatting, type ChartDataLabelFormattingEdit } from '../../../chart';
import type { PptxChartData, PptxChartDataLabelOptions } from '../types';
import type { NeutralChartEdit } from './chart-neutral-save-plan';
import { same } from './chart-neutral-save-common';
import { neutralTextEdit } from './chart-neutral-save-text';
import { neutralShapeEdit } from './chart-neutral-save-shape';
import { planNeutralChartLabels } from './chart-neutral-save-paint-labels';

function formatting(
	old: PptxChartDataLabelOptions,
	next: PptxChartDataLabelOptions,
	edits: NeutralChartEdit[],
	series?: number,
): void {
	const edit: ChartDataLabelFormattingEdit = {};
	if (!same(old.txPr, next.txPr)) {
		edit.font = neutralTextEdit(old.txPr, next.txPr);
		next.txPr = old.txPr;
	}
	if (!same(old.labelShape, next.labelShape)) {
		edit.shape = neutralShapeEdit(old.labelShape, next.labelShape);
		next.labelShape = old.labelShape;
	}
	if (old.showLeaderLines !== next.showLeaderLines) {
		edit.showLeaderLines = next.showLeaderLines ?? null;
		next.showLeaderLines = old.showLeaderLines;
	}
	if (!same(old.leaderLineStyle, next.leaderLineStyle)) {
		edit.leaderLine = next.leaderLineStyle
			? neutralShapeEdit(old.leaderLineStyle, next.leaderLineStyle)
			: null;
		next.leaderLineStyle = old.leaderLineStyle;
	}
	if (Object.keys(edit).length)
		edits.push((space) =>
			setChartDataLabelFormatting(space, edit, series === undefined ? {} : { series }),
		);
}

export function planNeutralLabelFormatting(
	before: PptxChartData,
	probe: PptxChartData,
	edits: NeutralChartEdit[],
): void {
	if (probe.style?.dataLabels)
		formatting(before.style?.dataLabels ?? {}, probe.style.dataLabels, edits);
	for (const [index, series] of probe.series.entries()) {
		const old = before.series[index];
		if (!old || same(old.dataLabelOptions, series.dataLabelOptions)) continue;
		const next = series.dataLabelOptions;
		if (!next) continue;
		formatting(old.dataLabelOptions ?? {}, next, edits, index);
		// Reuse the same option diff; the core target selects the series rather than chart groups.
		planNeutralChartLabels(before, probe, edits, index);
	}
}
