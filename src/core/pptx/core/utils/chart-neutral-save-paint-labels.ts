import {
	setChartDataLabels,
	setChartDataLabelsRange,
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
	seriesIndex?: number,
): void {
	const was =
		seriesIndex === undefined
			? (before.style ?? {})
			: { dataLabels: before.series[seriesIndex]?.dataLabelOptions };
	const now =
		seriesIndex === undefined
			? probe.style
			: { dataLabels: probe.series[seriesIndex]?.dataLabelOptions };
	const target = seriesIndex === undefined ? {} : { series: seriesIndex };
	if (!now) return;
	if (now.hasDataLabels === false && was.hasDataLabels !== false) {
		edits.push((space) =>
			setChartDataLabels(space, { show: false, preserveHiddenContent: true }, target),
		);
		if (was.hasDataLabels === undefined) delete now.hasDataLabels;
		else now.hasDataLabels = was.hasDataLabels;
		return;
	}
	const edit: ChartDataLabelsEdit = {};
	if (now.hasDataLabels === true && was.hasDataLabels !== true) edit.show = true;
	const previous = was.dataLabels ?? {};
	const labels = { ...now.dataLabels };
	if (
		JSON.stringify(labels.dataLabelsRange) !== JSON.stringify(previous.dataLabelsRange) ||
		labels.showDataLabelsRange !== previous.showDataLabelsRange
	) {
		const range = labels.dataLabelsRange;
		const shown = labels.showDataLabelsRange !== false;
		edits.push((space) => setChartDataLabelsRange(space, range ?? null, shown, target));
		for (const key of ['dataLabelsRange', 'showDataLabelsRange'] as const) {
			if (previous[key] === undefined) delete labels[key];
			else Object.assign(labels, { [key]: previous[key] });
		}
	}
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
	if (Object.keys(edit).length) edits.push((space) => setChartDataLabels(space, edit, target));
	if (was.hasDataLabels === undefined) delete now.hasDataLabels;
	else now.hasDataLabels = was.hasDataLabels;
	if (was.dataLabels === undefined && !Object.keys(labels).length) delete now.dataLabels;
	else now.dataLabels = labels;
	if (seriesIndex !== undefined) probe.series[seriesIndex]!.dataLabelOptions = now.dataLabels;
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
