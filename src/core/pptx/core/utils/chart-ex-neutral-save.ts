/**
 * Saves an edited ChartEx (`cx:chartSpace`) part through the shared chart core
 * (docs/agnostic-core-plan.md, chart step C5), the ChartEx counterpart of
 * `chart-neutral-save.ts`: the difference between the chart as loaded
 * (`before`) and as edited (`after`) becomes `chart/edit` chartex operations,
 * applied to the part as stored (`parseChartExSpace`) and written with
 * `writeChartExSpace`, so everything the edit does not touch stays as written.
 * Claims title, legend, labels, series colors/data, family options and supported layout changes.
 * Hierarchical category edits retain parent levels. Histogram and box-whisker series can grow.
 * Unmapped view changes retain the compatibility update; the package adapter separately refuses
 * unsupported sources and unsafe workbook writes before publishing either changed part.
 *
 * @module utils/chart-ex-neutral-save
 */

import {
	chartExWriteOptionsOf,
	parseChartExSpace,
	renameChartExSeries,
	setChartExCategories,
	setChartExDataLabelsVisible,
	setChartExLegend,
	setChartExSeriesValues,
	setChartExTitle,
	writeChartExSpace,
	type ChartExSpace,
} from '../../../chart/index';
import type { PptxChartData, PptxChartStyle } from '../types';
import { planChartExStyles } from './chart-ex-neutral-save-styles';
import { planChartExAxes } from './chart-ex-neutral-save-axes';
import { planChartExPointLabels } from './chart-ex-neutral-save-point-labels';
import { planChartExLabelOptions } from './chart-ex-neutral-save-label-options';
import {
	planChartExFamilyEdits,
	planChartExSeriesStructure,
} from './chart-ex-neutral-save-families';

export type ChartExEdit = (space: ChartExSpace) => ChartExSpace;

/** JSON with sorted keys and `undefined` members left out. */
function canonical(value: unknown): unknown {
	if (Array.isArray(value)) return value.map((item) => canonical(item) ?? null);
	if (value && typeof value === 'object')
		return Object.fromEntries(
			Object.keys(value)
				.sort()
				.flatMap((key) => {
					const item = canonical((value as Record<string, unknown>)[key]);
					return item === undefined ? [] : [[key, item]];
				}),
		);
	return value;
}

const same = (a: unknown, b: unknown): boolean =>
	JSON.stringify(canonical(a)) === JSON.stringify(canonical(b));

function revertStyle(
	probe: PptxChartData,
	before: PptxChartData,
	keys: readonly (keyof PptxChartStyle)[],
) {
	const style: Record<string, unknown> = { ...probe.style };
	for (const key of keys) {
		const value = before.style?.[key];
		if (value === undefined) delete style[key];
		else style[key] = value;
	}
	probe.style = style as PptxChartStyle;
}

function planTitle(before: PptxChartData, probe: PptxChartData, edits: ChartExEdit[]): void {
	const title = probe.title;
	const hasTitle = probe.style?.hasTitle;
	if (title === before.title && hasTitle === before.style?.hasTitle) return;
	if (hasTitle === false || (title === '' && hasTitle !== true))
		edits.push((space) => setChartExTitle(space, undefined));
	else if (title !== undefined && title !== before.title)
		edits.push((space) => setChartExTitle(space, title));
	else return;
	probe.title = before.title;
	revertStyle(probe, before, ['hasTitle']);
}

function planLegend(before: PptxChartData, probe: PptxChartData, edits: ChartExEdit[]): void {
	const was = before.style ?? {};
	const now = probe.style ?? {};
	if (was.hasLegend === now.hasLegend && was.legendPosition === now.legendPosition) return;
	// Without `hasLegend` the PowerPoint update leaves the legend alone, and so does this plan.
	if (now.hasLegend === false) edits.push((space) => setChartExLegend(space, { visible: false }));
	else if (now.hasLegend === true) {
		const position = now.legendPosition;
		edits.push((space) =>
			setChartExLegend(space, { visible: true, ...(position ? { position } : {}) }),
		);
	}
	revertStyle(probe, before, ['hasLegend', 'legendPosition']);
}

function planLabels(before: PptxChartData, probe: PptxChartData, edits: ChartExEdit[]): void {
	const shown = probe.style?.hasDataLabels;
	if (shown !== before.style?.hasDataLabels) {
		if (shown !== undefined) edits.push((space) => setChartExDataLabelsVisible(space, shown));
		revertStyle(probe, before, ['hasDataLabels']);
	}
}

const HIERARCHY = new Set(['treemap', 'sunburst', 'regionMap']);

function planSeries(before: PptxChartData, probe: PptxChartData, edits: ChartExEdit[]): void {
	if (!planChartExSeriesStructure(before, probe, edits)) return;
	probe.series.forEach((series, i) => {
		const previous = before.series[i];
		if (!previous) return;
		if (series.name !== previous.name) {
			const name = series.name;
			edits.push((space) => renameChartExSeries(space, i, name));
			series.name = previous.name;
		}
		if (!same(series.values, previous.values) || !same(series.blanks, previous.blanks)) {
			const values = series.values.map((value, point) => (series.blanks?.[point] ? null : value));
			edits.push((space) => setChartExSeriesValues(space, i, values));
			series.values = previous.values;
			if (previous.blanks === undefined) delete series.blanks;
			else series.blanks = previous.blanks;
		}
	});
	if (
		(!same(before.categories, probe.categories) ||
			!same(before.categoryLevels, probe.categoryLevels)) &&
		before.chartType !== 'regionMap'
	) {
		const categories = HIERARCHY.has(probe.chartType)
			? same(before.categoryLevels, probe.categoryLevels)
				? [probe.categories, ...(probe.categoryLevels?.slice(1) ?? [])]
				: (probe.categoryLevels ?? [probe.categories])
			: [...probe.categories];
		edits.push((space) => setChartExCategories(space, categories));
		probe.categories = before.categories;
		probe.categoryLevels = before.categoryLevels;
	}
	probe.series = [
		...probe.series.slice(0, before.series.length),
		...before.series.slice(probe.series.length),
	];
}

/** The chartex edits that turn the part saved as `before` into `after`, or `undefined`. */
export function planChartExNeutralSave(
	before: PptxChartData,
	after: PptxChartData,
): ChartExEdit[] | undefined {
	const probe = structuredClone(after);
	const edits: ChartExEdit[] = [];
	planTitle(before, probe, edits);
	planLegend(before, probe, edits);
	planLabels(before, probe, edits);
	planChartExStyles(before, probe, edits);
	planChartExLabelOptions(before, probe, edits);
	planChartExAxes(before, probe, edits);
	planChartExPointLabels(before, probe, edits);
	planChartExFamilyEdits(before, probe, edits);
	planSeries(before, probe, edits);
	return same(probe, before) ? edits : undefined;
}

/**
 * `xml` (the ChartEx part as stored, which `before` describes) saved with the edits that make it
 * `after`, or `undefined` when the shared core cannot save this part or this edit yet.
 */
export function saveChartExPartThroughNeutral(
	xml: string,
	before: PptxChartData,
	after: PptxChartData,
): string | undefined {
	const edits = planChartExNeutralSave(before, after);
	if (!edits) return undefined;
	try {
		const parsed = parseChartExSpace(xml);
		if (parsed.issues.length > 0) return undefined;
		let space = parsed.chartExSpace;
		for (const edit of edits) space = edit(space);
		return writeChartExSpace(space, chartExWriteOptionsOf(xml));
	} catch {
		return undefined;
	}
}
