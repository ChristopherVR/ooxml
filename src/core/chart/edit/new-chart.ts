// A new chart part for a product's Insert Chart: the data sheet laid out in the standard layout of
// the embedded workbook (`workbook/layout.ts`), one clustered column group with Excel's two
// default axes (`type-axes.ts`), then switched to the picked type through `setChartType`, so a new
// chart of any family is built by the same rules as a chart changed to it. Office's defaults: the
// 4 x 3 sample Word and PowerPoint open with, a title and a legend at the bottom.
//
// A host with defaults of its own (a chart written with no workbook, its own gap width, no
// gridlines) asks for literal data and the `minimal` profile (`new-chart-minimal.ts`), then states
// what it wants through `group` and the edit operations.
import type { ChartDataSheet } from '../data-sheet';
import type { ChartPlotGroup, ChartSpace } from '../model';
import type { ChartSpaceSeries } from '../model-series';
import { DEFAULT_CHART_SHEET_NAME, standardChartFormulas } from '../workbook/layout';
import { editChartSpace } from './draft';
import { assertChartGroupOption, isChartGroupOption } from './group-options';
import { setChartLegend, type ChartLegendPosition } from './legend-labels';
import { minimalChartAxes, minimalChartGroup, minimalChartTitle } from './new-chart-minimal';
import { categorySource, numberSource } from './series-data';
import { setChartTitle } from './title';
import { setChartType, type ChartTypeTarget } from './type';
import { DEFAULT_CHART_AXIS_IDS, defaultChartAxes } from './type-axes';
import { GROUP_FIELDS } from './type-schema';

/** The sample data Word and PowerPoint put behind a new chart. */
export const DEFAULT_CHART_DATA_SHEET: ChartDataSheet = {
	categories: ['Category 1', 'Category 2', 'Category 3', 'Category 4'],
	series: [
		{ name: 'Series 1', values: [4.3, 2.5, 3.5, 4.5] },
		{ name: 'Series 2', values: [2.4, 4.4, 1.8, 2.8] },
		{ name: 'Series 3', values: [2, 2, 3, 5] },
	],
};

/** Group settings a new chart can start with (the type, direction and grouping are the target's). */
export type NewChartGroupSettings = Partial<
	Omit<
		ChartPlotGroup,
		'kind' | 'element' | 'is3D' | 'series' | 'axisIds' | 'barDirection' | 'grouping'
	>
>;

export interface NewChartOptions {
	/** The chart title; absent for no title (`c:autoTitleDeleted`). */
	title?: string;
	/** The worksheet the formulas name; Office's `Sheet1` by default. */
	sheetName?: string;
	/**
	 * Where the series read their data. `workbook` (the default): the cells of the embedded
	 * workbook's standard layout, with cached values. `literal`: the values are written into the
	 * part (`c:v` series names, `c:strLit` or `c:numLit` data), for a chart with no workbook.
	 */
	data?: 'workbook' | 'literal';
	/** Where the legend sits, `null` for no legend; Office's `b` by default. */
	legendPosition?: ChartLegendPosition | null;
	/**
	 * `office` (the default): what Word and PowerPoint write for a new chart (square corners,
	 * Excel's axes with value gridlines and tick marks, a 219% gap and -27% overlap, a title under
	 * Excel's paragraph properties). `minimal`: only what the schema requires for the type, plus
	 * `c:varyColors` and `c:delete`, and a title of one bare run.
	 */
	profile?: 'office' | 'minimal';
	/**
	 * Settings of the new group, which its element must allow (`type-schema.ts`); the gap width,
	 * overlap, first slice angle and hole size must be whole numbers in their schema ranges.
	 */
	group?: NewChartGroupSettings;
	/**
	 * The ids of the category (or X value), value (or Y value) and series axes, in that order;
	 * Excel's `500000001` and `500000002` by default (and `500000003` for a series axis).
	 */
	axisIds?: readonly number[];
	/**
	 * Whether the category axis is a date axis (`c:dateAx`), for categories that are dates; a type
	 * whose first axis is not a category axis (scatter, bubble, the pie family) ignores it.
	 */
	dateAxis?: boolean;
}

/** Fields a category axis has and a date axis does not (CT_CatAx against CT_DateAx). */
const CATEGORY_ONLY = [
	'labelAlign',
	'noMultiLevelLabels',
	'tickLabelSkip',
	'tickMarkSkip',
] as const;

/** `space` with its category axis turned into a date axis. */
function withDateAxis(space: ChartSpace): ChartSpace {
	return editChartSpace(space, (draft) => {
		const axis = draft.plotArea.axes[0];
		if (axis?.kind !== 'cat') return;
		axis.kind = 'date';
		for (const key of CATEGORY_ONLY) delete axis[key];
	});
}

function series(
	sheet: ChartDataSheet,
	index: number,
	options: NewChartOptions,
	sheetName: string,
): ChartSpaceSeries {
	const entry = sheet.series[index]!;
	if (options.data === 'literal')
		return {
			index,
			order: index,
			tx: { value: entry.name, text: entry.name },
			...(options.profile === 'minimal' ? {} : { invertIfNegative: false }),
			dataPoints: [],
			categories: categorySource(sheet.categories, undefined),
			values: numberSource(entry.values, undefined),
		};
	const formulas = standardChartFormulas(sheetName, index, sheet.categories.length);
	return {
		index,
		order: index,
		tx: { reference: categorySource([entry.name], undefined, formulas.name), text: entry.name },
		...(options.profile === 'minimal' ? {} : { invertIfNegative: false }),
		dataPoints: [],
		categories: categorySource(sheet.categories, undefined, formulas.categories),
		values: numberSource(entry.values, undefined, formulas.values),
	};
}

function axisIds(options: NewChartOptions): [number, number, number] {
	const ids = options.axisIds ?? [];
	const [first, second, third] = [
		ids[0] ?? DEFAULT_CHART_AXIS_IDS[0],
		ids[1] ?? DEFAULT_CHART_AXIS_IDS[1],
		ids[2] ?? DEFAULT_CHART_AXIS_IDS[1] + 1,
	];
	for (const id of [first, second, third])
		if (!Number.isSafeInteger(id) || id < 0)
			throw new RangeError('Chart axis ids must be non-negative whole numbers');
	if (new Set([first, second, third]).size !== 3)
		throw new RangeError('Chart axis ids must differ');
	return [first, second, third];
}

/** `space` with its axes renumbered to `ids`, in plot-area order. */
function withAxisIds(space: ChartSpace, ids: readonly number[]): ChartSpace {
	return editChartSpace(space, (draft) => {
		const map = new Map<number, number>();
		draft.plotArea.axes.forEach((axis, position) => {
			if (axis.id !== undefined && ids[position] !== undefined) map.set(axis.id, ids[position]);
		});
		const renumber = (id: number) => map.get(id) ?? id;
		for (const axis of draft.plotArea.axes) {
			if (axis.id !== undefined) axis.id = renumber(axis.id);
			if (axis.crossAxisId !== undefined) axis.crossAxisId = renumber(axis.crossAxisId);
		}
		for (const group of draft.plotArea.groups) group.axisIds = group.axisIds.map(renumber);
	});
}

function withGroupSettings(space: ChartSpace, settings: NewChartGroupSettings): ChartSpace {
	return editChartSpace(space, (draft) => {
		const group = draft.plotArea.groups[0]!;
		const allowed: readonly string[] = GROUP_FIELDS[group.element] ?? [];
		for (const [key, value] of Object.entries(settings)) {
			if (value === undefined) continue;
			if (!allowed.includes(key)) throw new RangeError(`A ${group.element} has no ${key} setting`);
			if (isChartGroupOption(key)) assertChartGroupOption(key, value as number);
			Object.assign(group, { [key]: structuredClone(value) });
		}
	});
}

function officeChartSpace(
	sheet: ChartDataSheet,
	target: ChartTypeTarget | undefined,
	options: NewChartOptions,
	sheetName: string,
): ChartSpace {
	const group: ChartPlotGroup = {
		kind: 'bar',
		element: 'barChart',
		is3D: false,
		barDirection: 'col',
		grouping: 'clustered',
		varyColors: false,
		gapWidth: 219,
		overlap: -27,
		series: sheet.series.map((_, index) => series(sheet, index, options, sheetName)),
		axisIds: [...DEFAULT_CHART_AXIS_IDS],
	};
	const space: ChartSpace = {
		roundedCorners: false,
		autoTitleDeleted: options.title === undefined,
		plotArea: { layout: {}, groups: [group], axes: defaultChartAxes(false, false) },
		plotVisibleOnly: true,
		displayBlanksAs: 'gap',
	};
	const column = !target || (target.kind === 'bar' && (target.barDirection ?? 'col') === 'col');
	if (target && !(column && !target.is3D && (target.grouping ?? 'clustered') === 'clustered'))
		return setChartType(space, target);
	return space;
}

function minimalChartSpace(
	sheet: ChartDataSheet,
	target: ChartTypeTarget | undefined,
	options: NewChartOptions,
	sheetName: string,
): ChartSpace {
	const group = minimalChartGroup(
		target ?? { kind: 'bar' },
		sheet.series.map((_, index) => series(sheet, index, options, sheetName)),
	);
	if (group.kind === 'stock' && ![3, 4].includes(group.series.length))
		throw new RangeError('A stock chart requires three or four price series');
	const axes = minimalChartAxes(group, axisIds(options));
	group.axisIds = axes.flatMap((axis) => (axis.id === undefined ? [] : [axis.id]));
	return {
		autoTitleDeleted: options.title === undefined,
		plotArea: { layout: {}, groups: [group], axes },
		plotVisibleOnly: true,
		displayBlanksAs: 'gap',
	};
}

/** A new chart part showing `sheet` as `target` (a column chart when `target` is absent). */
export function newChartSpace(
	sheet: ChartDataSheet = DEFAULT_CHART_DATA_SHEET,
	target?: ChartTypeTarget,
	options: NewChartOptions = {},
): ChartSpace {
	const sheetName = options.sheetName ?? DEFAULT_CHART_SHEET_NAME;
	let space =
		options.profile === 'minimal'
			? minimalChartSpace(sheet, target, options, sheetName)
			: officeChartSpace(sheet, target, options, sheetName);
	if (options.profile !== 'minimal' && options.axisIds !== undefined)
		space = withAxisIds(space, axisIds(options));
	if (options.dateAxis) space = withDateAxis(space);
	if (options.group) space = withGroupSettings(space, options.group);
	if (options.title !== undefined)
		space =
			options.profile === 'minimal'
				? { ...space, title: minimalChartTitle(options.title), autoTitleDeleted: false }
				: setChartTitle(space, options.title);
	const legend = options.legendPosition === undefined ? 'b' : options.legendPosition;
	return legend === null ? space : setChartLegend(space, { position: legend });
}
