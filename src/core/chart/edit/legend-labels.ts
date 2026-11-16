// Legend and data-label edits. Legend: PowerPoint's `setChartLegend` semantics (a position shows a
// hidden legend), written the way Excel adds one (`c:legendPos`, `c:overlay val="0"`). Data labels:
// PowerPoint's `setChartDataLabels` semantics (a content flag or position shows the labels; showing
// labels that say nothing turns the value on), with Office's explicit switches on new labels.
import type { ChartSpace } from '../model';
import type { ChartDataLabels } from '../model-series';
import { assertChartLabelPosition } from '../label-positions';
import { assignOrDelete, chartSeriesEntryAt, chartSpaceSeries, editChartSpace } from './draft';

/** ST_LegendPos. */
export type ChartLegendPosition = 'b' | 'tr' | 'l' | 'r' | 't';

export interface ChartLegendEdit {
	visible?: boolean;
	position?: ChartLegendPosition;
	overlay?: boolean;
}

/** Shows, hides, places or overlays the legend. Hiding removes `c:legend` and its entries. */
export function setChartLegend(space: ChartSpace, edit: ChartLegendEdit): ChartSpace {
	return editChartSpace(space, (draft) => {
		if (edit.visible === false) {
			delete draft.legend;
			return;
		}
		if (edit.visible || edit.position !== undefined)
			draft.legend ??= { position: edit.position ?? 'r', overlay: false, entries: [] };
		if (!draft.legend) return;
		if (edit.position !== undefined) draft.legend.position = edit.position;
		if (edit.overlay !== undefined) draft.legend.overlay = edit.overlay;
	});
}

export interface ChartDataLabelsEdit {
	show?: boolean;
	/** Hide with c:delete while retaining authored switches and series overrides. */
	preserveHiddenContent?: boolean;
	showValue?: boolean;
	showCategoryName?: boolean;
	showSeriesName?: boolean;
	showPercent?: boolean;
	showBubbleSize?: boolean;
	showLegendKey?: boolean;
	/** `c:dLblPos` (`ctr`, `inEnd`, `outEnd`, `bestFit`...); `null` returns to the default. */
	position?: string | null;
	/** `null` removes the separator. */
	separator?: string | null;
	/** A number format code not linked to the source; `null` removes it. */
	numberFormat?: string | null;
}

const CONTENT = [
	'showValue',
	'showCategoryName',
	'showSeriesName',
	'showPercent',
	'showBubbleSize',
] as const;
const SWITCHES = ['showLegendKey', ...CONTENT] as const;

/** New labels with every switch written, as Office writes them. */
const newLabels = (): ChartDataLabels => ({
	labels: [],
	showLegendKey: false,
	showValue: false,
	showCategoryName: false,
	showSeriesName: false,
	showPercent: false,
	showBubbleSize: false,
});

function applyLabels(current: ChartDataLabels | undefined, edit: ChartDataLabelsEdit) {
	const contentEdit =
		SWITCHES.some((key) => edit[key] !== undefined) ||
		edit.position !== undefined ||
		edit.separator !== undefined ||
		edit.numberFormat !== undefined;
	if (edit.show === false) {
		if (edit.preserveHiddenContent) return { ...(current ?? { labels: [] }), deleted: true };
		if (!current) return newLabels();
		for (const key of SWITCHES) current[key] = false;
		delete current.deleted;
		return current;
	}
	if (!edit.show && !contentEdit) return current;
	const labels = current ?? newLabels();
	delete labels.deleted;
	for (const key of SWITCHES) if (edit[key] !== undefined) labels[key] = edit[key];
	if (edit.show && edit.showValue === undefined && !CONTENT.some((key) => labels[key] === true))
		labels.showValue = true;
	if (edit.position !== undefined) assignOrDelete(labels, 'position', edit.position || undefined);
	if (edit.separator !== undefined)
		assignOrDelete(labels, 'separator', edit.separator ?? undefined);
	if (edit.numberFormat !== undefined)
		assignOrDelete(
			labels,
			'numberFormat',
			edit.numberFormat ? { formatCode: edit.numberFormat, sourceLinked: false } : undefined,
		);
	return labels;
}

/**
 * The labels a series takes when an edit only formats them (a number format or separator) and the
 * series has none of its own: the switches and position of its group's labels, so the series shows
 * what it showed before. Any other edit starts from Office's switched-off labels.
 */
function inheritedLabels(
	group: ChartDataLabels | undefined,
	edit: ChartDataLabelsEdit,
): ChartDataLabels | undefined {
	const formatOnly =
		edit.show === undefined &&
		edit.position === undefined &&
		SWITCHES.every((key) => edit[key] === undefined);
	if (!formatOnly || !group || group.deleted) return undefined;
	const labels = newLabels();
	for (const key of SWITCHES) if (group[key] !== undefined) labels[key] = group[key];
	if (group.position !== undefined) labels.position = group.position;
	if (group.separator !== undefined) labels.separator = group.separator;
	return labels;
}

/**
 * Edits the data labels of the whole chart (every group's `c:dLbls`) or of one series. Hiding the
 * chart's labels also removes every series' own labels, as Office's Chart Elements box does; hiding
 * one series' labels writes switched-off labels on it so the group's labels do not show through.
 * A position must be one every edited group accepts (`label-positions.ts`).
 */
export function setChartDataLabels(
	space: ChartSpace,
	edit: ChartDataLabelsEdit,
	target: { series?: number } = {},
): ChartSpace {
	return editChartSpace(space, (draft) => {
		if (target.series !== undefined) {
			const { group, series } = chartSeriesEntryAt(draft, target.series);
			if (edit.position) assertChartLabelPosition(group, edit.position);
			const next = applyLabels(series.dataLabels ?? inheritedLabels(group.dataLabels, edit), edit);
			assignOrDelete(series, 'dataLabels', next);
			return;
		}
		if (edit.position)
			for (const group of draft.plotArea.groups)
				if (group.kind !== 'surface') assertChartLabelPosition(group, edit.position);
		for (const group of draft.plotArea.groups) {
			if (group.kind === 'surface') continue;
			if (edit.show === false && !edit.preserveHiddenContent && !group.dataLabels) continue;
			assignOrDelete(group, 'dataLabels', applyLabels(group.dataLabels, edit));
		}
		if (edit.show === false && !edit.preserveHiddenContent)
			for (const series of chartSpaceSeries(draft)) delete series.dataLabels;
	});
}
