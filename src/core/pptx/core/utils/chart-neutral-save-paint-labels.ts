import {
	setChartDataLabels,
	setChartSeriesColor,
	type ChartDataLabelsEdit,
} from '../../../chart/index';
import type { PptxChartData } from '../types';
import type { NeutralChartEdit } from './chart-neutral-save-plan';

const FLAGS = {
	showValue: 'showValue',
	showCategory: 'showCategoryName',
	showSeriesName: 'showSeriesName',
	showPercent: 'showPercent',
	showBubbleSize: 'showBubbleSize',
	showLegendKey: 'showLegendKey',
} as const;

/** Claim supported label changes without flattening unrelated group or series formatting. */
export function planNeutralChartLabels(
	before: PptxChartData,
	probe: PptxChartData,
	edits: NeutralChartEdit[],
): void {
	const was = before.style ?? {};
	const now = probe.style;
	if (!now || now.hasDataLabels === false) return;
	const edit: ChartDataLabelsEdit = {};
	if (now.hasDataLabels === true && was.hasDataLabels !== true) edit.show = true;
	const previous = was.dataLabels ?? {};
	const labels = { ...now.dataLabels };
	for (const [source, target] of Object.entries(FLAGS) as [
		keyof typeof FLAGS,
		(typeof FLAGS)[keyof typeof FLAGS],
	][]) {
		if (labels[source] === previous[source]) continue;
		// Removing a flag returns to inheritance, unlike explicitly switching it off.
		if (labels[source] === undefined) continue;
		edit[target] = labels[source];
		if (previous[source] === undefined) delete labels[source];
		else labels[source] = previous[source];
	}
	for (const key of ['position', 'separator', 'numberFormat'] as const) {
		if (labels[key] === previous[key]) continue;
		edit[key] = labels[key] ?? null;
		if (previous[key] === undefined) delete labels[key];
		else Object.assign(labels, { [key]: previous[key] });
	}
	if (!Object.keys(edit).length) return;
	edits.push((space) => setChartDataLabels(space, edit));
	if (was.hasDataLabels === undefined) delete now.hasDataLabels;
	else now.hasDataLabels = was.hasDataLabels;
	if (was.dataLabels === undefined && !Object.keys(labels).length) delete now.dataLabels;
	else now.dataLabels = labels;
}

/** Solid colour edits share the neutral family's line/fill and transparency policy. */
export function planNeutralChartSeriesColors(
	before: PptxChartData,
	probe: PptxChartData,
	edits: NeutralChartEdit[],
): void {
	if (before.series.length !== probe.series.length) return;
	probe.series.forEach((series, index) => {
		const previous = before.series[index]!;
		if (series.color === previous.color && series.gradientFill === previous.gradientFill) return;
		if (series.gradientFill !== undefined) return;
		const value = series.color?.replace(/^#/, '');
		if (value !== undefined && !/^[\da-f]{6}$/i.test(value)) return;
		edits.push((space) =>
			setChartSeriesColor(space, index, value ? { kind: 'srgb', value, transforms: [] } : null),
		);
		if (previous.color === undefined) delete series.color;
		else series.color = previous.color;
		if (previous.gradientFill === undefined) delete series.gradientFill;
		else series.gradientFill = previous.gradientFill;
	});
}
