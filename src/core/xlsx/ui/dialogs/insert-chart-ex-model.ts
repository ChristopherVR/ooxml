// Insert Chart over the selection for Excel's extended families (waterfall, histogram, pareto, box
// and whisker, treemap, sunburst, funnel and region map). The chart reads its data from the
// selected cells, as Excel's own does: every dimension's formula is a direct range on the sheet,
// with the current values cached. Leading text columns of a treemap or sunburst selection are
// its hierarchy levels; a region map's categories are place names.
import {
	applyChartExOfficeDefaults,
	bindChartExReferences,
	chartExSummaryFromModel,
	createChartExSpace,
	type ChartExFamily,
	type ChartExSpace,
} from '../../../chart';
import {
	formatAddress,
	getCell,
	normalizeRange,
	quoteSheetName,
	type CellRange,
	type UnsupportedDrawingObject,
	type Workbook,
} from '../../index';
import { buildChart } from './insert-chart-model';

const absolute = (start: { row: number; col: number }, end: { row: number; col: number }) => {
	const a = (row: number, col: number) =>
		formatAddress({ row, col }).replace(/^([A-Z]+)(\d+)$/, '$$$1$$$2');
	const from = a(start.row, start.col);
	const to = a(end.row, end.col);
	return from === to ? from : `${from}:${to}`;
};

/** The extended families Insert Chart offers after the classic types, with their English names. */
export const INSERT_CHART_EX_TYPES: ReadonlyArray<readonly [family: ChartExFamily, label: string]> =
	[
		['waterfall', 'Waterfall'],
		['clusteredColumn', 'Histogram'],
		['pareto', 'Pareto'],
		['boxWhisker', 'Box and Whisker'],
		['treemap', 'Treemap'],
		['sunburst', 'Sunburst'],
		['funnel', 'Funnel'],
		['regionMap', 'Map'],
	];

/** The families that draw several series; the others draw the first one. */
const MULTI_SERIES = new Set<ChartExFamily>(['boxWhisker', 'clusteredColumn']);
const HIERARCHY = new Set<ChartExFamily>(['treemap', 'sunburst']);

/** A new extended chart object over `range`; add it with the session's `addChartEx`. */
export function buildChartEx(
	workbook: Workbook,
	sheetIndex: number,
	range: CellRange,
	family: ChartExFamily,
	options: { title?: string } = {},
): UnsupportedDrawingObject & { chartExSpace: ChartExSpace } {
	const ws = workbook.sheets[sheetIndex];
	if (!ws) throw new RangeError(`No sheet at index ${sheetIndex}`);
	const r = normalizeRange(range);
	const prefix = `${quoteSheetName(ws.name)}!`;
	const text = (row: number, col: number) => {
		const value = getCell(ws, row, col)?.value;
		return typeof value === 'string' && value !== '';
	};
	const header = [...Array(r.end.col - r.start.col + 1).keys()].some((c) =>
		text(r.start.row, r.start.col + c),
	);
	const first = r.start.row + (header ? 1 : 0);
	// Leading text columns (any text in the data rows) are category levels.
	let levels = 0;
	while (
		r.start.col + levels < r.end.col &&
		[...Array(r.end.row - first + 1).keys()].some((i) => text(first + i, r.start.col + levels))
	)
		levels++;
	if (!HIERARCHY.has(family)) levels = Math.min(levels, 1);
	const valueRange: CellRange =
		levels > 1 ? { start: { row: r.start.row, col: r.start.col + levels - 1 }, end: r.end } : r;
	const chart = buildChart(workbook, sheetIndex, valueRange, 'column');
	const series = MULTI_SERIES.has(family) ? chart.series : chart.series.slice(0, 1);
	if (!series.length || series.every((s) => s.values.every((v) => v === null)))
		throw new RangeError('Select numbers to chart.');
	const count = Math.max(...series.map((s) => s.values.length));
	const column = (col: number) =>
		Array.from({ length: count }, (_, i) => {
			const value = getCell(ws, first + i, col)?.value;
			return value === null || value === undefined ? '' : String(value);
		});
	// Office stores hierarchy levels leaf first; worksheet columns are outermost first.
	const categoryLevels =
		levels > 0
			? Array.from({ length: levels }, (_, l) => column(r.start.col + levels - 1 - l))
			: [Array.from({ length: count }, (_, i) => String(i + 1))];
	const space = createChartExSpace(
		{
			categories: categoryLevels[0]!,
			categoryLevels,
			series: series.map((s, i) => ({ name: s.name ?? `Series ${i + 1}`, values: s.values })),
		},
		family,
		options.title ? { title: options.title } : {},
	);
	const categories =
		levels > 0
			? prefix +
				absolute(
					{ row: first, col: r.start.col },
					{ row: first + count - 1, col: r.start.col + levels - 1 },
				)
			: undefined;
	const bound = bindChartExReferences(applyChartExOfficeDefaults(space, family), {
		...(categories ? { categories } : {}),
		series: series.map((s) => ({
			...(s.nameRef ? { name: s.nameRef } : {}),
			...(s.valuesRef ? { values: s.valuesRef } : {}),
		})),
	});
	const summary = chartExSummaryFromModel(bound);
	const object: UnsupportedDrawingObject & { chartExSpace: ChartExSpace } = {
		kind: 'unsupported',
		anchor: chart.anchor,
		description: `chart (${summary.layoutId})`,
		chartExSpace: bound,
	};
	if (summary.summary) object.chartEx = { ...summary.summary, kind: 'chart', anchor: chart.anchor };
	return object;
}
