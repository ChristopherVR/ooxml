// Series edits: add, remove, reorder, rename, values and categories (single points and category
// rows are in `series-rows.ts`). Semantics follow PowerPoint's chart operations (`addChartSeries`,
// `removeChartSeries`, `setChartCategories`, `updateChartSeriesValues`): the category list is
// shared, so a category edit applies to every series unless one is named.
// Scatter and bubble series hold their data as x and y values; the operations map categories to
// x values and values to y values there.
import type { ChartSpace } from '../model';
import type { ChartShapeProperties, ChartSpaceSeries } from '../model-series';
import { chartSeriesEntries, chartSeriesEntryAt, chartSpaceSeries, editChartSpace } from './draft';
import type { ChartSeriesValue } from '../data-sheet';
import { categorySource, numberSource, type ChartCategoryValue } from './series-data';
import { chartCategoryKey, chartValueKey } from './series-keys';

export { isXYChartKind } from './series-keys';

export interface NewChartSeries {
	name?: string;
	/** A reference for the name (`Sheet1!$B$1`), cached as `name`. */
	nameFormula?: string;
	values: readonly ChartSeriesValue[];
	valuesFormula?: string;
	/** Defaults to the categories of the chart's first series. */
	categories?: readonly ChartCategoryValue[];
	categoriesFormula?: string;
	spPr?: ChartShapeProperties;
	/** The group to add to; defaults to the last group. */
	groupIndex?: number;
}

const nextNumber = (numbers: (number | undefined)[]) =>
	numbers.reduce<number>((max, value) => (value !== undefined && value > max ? value : max), -1) +
	1;

/** Appends a series; its `c:idx` and `c:order` follow the highest in the chart. */
export function addChartSeries(space: ChartSpace, input: NewChartSeries): ChartSpace {
	return editChartSpace(space, (draft) => {
		const groups = draft.plotArea.groups;
		const group = groups.find((_, i) => i === (input.groupIndex ?? groups.length - 1));
		if (!group) throw new RangeError('The chart has no chart group to add a series to');
		const all = chartSpaceSeries(draft);
		const template = all[0];
		const series: ChartSpaceSeries = {
			index: nextNumber(all.map((item) => item.index)),
			order: nextNumber(all.map((item) => item.order)),
			dataPoints: [],
		};
		if (input.nameFormula !== undefined)
			series.tx = {
				reference: categorySource(
					input.name === undefined ? [] : [input.name],
					undefined,
					input.nameFormula,
				),
				...(input.name !== undefined ? { text: input.name } : {}),
			};
		else if (input.name !== undefined) series.tx = { value: input.name, text: input.name };
		if (input.spPr) series.spPr = structuredClone(input.spPr);
		const templateCategories = template?.categories ?? template?.xValues;
		if (input.categories !== undefined)
			series[chartCategoryKey(group)] = categorySource(
				input.categories,
				undefined,
				input.categoriesFormula,
			);
		else if (templateCategories)
			series[chartCategoryKey(group)] = structuredClone(templateCategories);
		series[chartValueKey(group)] = numberSource(input.values, undefined, input.valuesFormula);
		group.series.push(series);
	});
}

/**
 * Removes the series at a flattened index. The plot order (`c:order`) of every later series moves
 * up one, as Excel renumbers it; `c:idx`, which formatting and legend entries refer to, is kept.
 */
export function removeChartSeries(space: ChartSpace, index: number): ChartSpace {
	return editChartSpace(space, (draft) => {
		const { group, position } = chartSeriesEntryAt(draft, index);
		const [removed] = group.series.splice(position, 1);
		const order = removed?.order;
		if (order === undefined) return;
		for (const { series } of chartSeriesEntries(draft))
			if (series.order !== undefined && series.order > order) series.order -= 1;
	});
}

/**
 * Moves a series to another position of its own group (moving across groups would change its
 * chart type). The group's `c:order` values are reassigned in the new order, so the plot order
 * follows; `c:idx`, which formatting and legend entries refer to, is kept.
 */
export function moveChartSeries(space: ChartSpace, from: number, to: number): ChartSpace {
	return editChartSpace(space, (draft) => {
		const source = chartSeriesEntryAt(draft, from);
		const target = chartSeriesEntryAt(draft, to);
		if (source.group !== target.group)
			throw new RangeError('A series can only move within its own chart group');
		const list = source.group.series;
		const orders = list.map((item, i) => item.order ?? i).sort((a, b) => a - b);
		const [moved] = list.splice(source.position, 1);
		list.splice(target.position, 0, moved!);
		list.forEach((item, i) => (item.order = orders[i]!));
	});
}

export interface RenameChartSeriesOptions {
	/**
	 * Keep a name taken from a cell and refresh its cached text, as PowerPoint's and Word's data
	 * sheet does when the cell itself is retyped (the embedded workbook is the caller's to update).
	 */
	keepReference?: boolean;
}

/**
 * Renames a series. By default a name taken from a cell becomes a literal name, as typing it into
 * Excel's series name box does; `keepReference` follows the data sheet instead.
 */
export function renameChartSeries(
	space: ChartSpace,
	index: number,
	name: string,
	options: RenameChartSeriesOptions = {},
): ChartSpace {
	return editChartSpace(space, (draft) => {
		const { series } = chartSeriesEntryAt(draft, index);
		const reference = series.tx?.reference;
		if (options.keepReference && reference?.kind === 'strRef' && reference.formula) {
			const cache = { type: 'string' as const, pointCount: 1, points: [{ index: 0, value: name }] };
			series.tx = { reference: { ...reference, cache }, text: name };
			return;
		}
		series.tx = { value: name, text: name };
	});
}

/** Replaces a series' values (y values of scatter and bubble series). */
export function setChartSeriesValues(
	space: ChartSpace,
	index: number,
	values: readonly ChartSeriesValue[],
	options: { formula?: string } = {},
): ChartSpace {
	return editChartSpace(space, (draft) => {
		const { group, series } = chartSeriesEntryAt(draft, index);
		const key = chartValueKey(group);
		series[key] = numberSource(values, series[key], options.formula);
	});
}

/** Replaces the categories (x values of scatter and bubble series) of every series, or one. */
export function setChartCategories(
	space: ChartSpace,
	categories: readonly ChartCategoryValue[],
	options: { series?: number; formula?: string } = {},
): ChartSpace {
	return editChartSpace(space, (draft) => {
		const entries =
			options.series === undefined
				? draft.plotArea.groups.flatMap((group) =>
						group.series.map((series) => ({ group, series })),
					)
				: [chartSeriesEntryAt(draft, options.series)];
		for (const { group, series } of entries) {
			const key = chartCategoryKey(group);
			series[key] = categorySource(categories, series[key], options.formula);
		}
	});
}
