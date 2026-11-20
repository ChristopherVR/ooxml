// Map ChartEx caches to exact local workbook vectors, including multi-column hierarchies.
import type {
	ChartExData,
	ChartExDimension,
	ChartExLevel,
	ChartExSpace,
	ChartExText,
} from './model';
import { formatCellAddress, parseChartFormulaRange } from '../workbook/cellref';
import { quoteChartSheetName } from '../workbook/layout';
import { chartCacheCellChanges, chartCacheCellCount } from '../workbook/space-writes';
import type { ChartWorkbookWrite } from '../workbook/writer';

export type ChartExFormulaResolver = (formula: string) => string | undefined;
const absolute = (col: number, row: number) =>
	formatCellAddress({ col, row }).replace(/^([A-Z]+)(\d+)$/u, '$$$1$$$2');

/** A direct local range, optionally resolved from a host's defined name. External refs are refused. */
export function chartExLocalRange(formula: string, resolve?: ChartExFormulaResolver) {
	const direct = parseChartFormulaRange(formula) ? formula : resolve?.(formula);
	const range = direct ? parseChartFormulaRange(direct.replace(/^=/, '')) : undefined;
	if (
		!range ||
		/[\[\]:]/.test(range.sheetName) ||
		range.start.row < 0 ||
		range.start.col < 0 ||
		range.end.row < range.start.row ||
		range.end.col < range.start.col ||
		range.end.row > 1048575 ||
		range.end.col > 16383
	)
		throw new Error('ChartEx data editing requires direct local workbook ranges.');
	return range;
}

interface LevelRange {
	formula: string;
	isNumeric: boolean;
	level: ChartExLevel;
	cellCount: number;
}

/** The worksheet vector of every cached level of a dimension. */
function dimensionRanges(
	dimension: ChartExDimension,
	resolve?: ChartExFormulaResolver,
): LevelRange[] {
	if (!dimension.formula) return [];
	const range = chartExLocalRange(dimension.formula.formula, resolve);
	const horizontal =
		dimension.formula.direction === 'row' ||
		(range.start.row === range.end.row && dimension.levels.length === 1);
	const levelCount = horizontal
		? range.end.row - range.start.row + 1
		: range.end.col - range.start.col + 1;
	const count = horizontal
		? range.end.col - range.start.col + 1
		: range.end.row - range.start.row + 1;
	if (levelCount !== dimension.levels.length)
		throw new Error('ChartEx source dimensions do not match cached levels.');
	return dimension.levels.map((level, index) => {
		if (chartCacheCellCount(level) !== count)
			throw new Error('ChartEx source range does not match cached points.');
		// Office stores hierarchy levels leaf-first, while worksheet columns are outermost-first.
		const offset = dimension.kind === 'strDim' ? levelCount - 1 - index : index;
		const start = {
			col: range.start.col + (horizontal ? 0 : offset),
			row: range.start.row + (horizontal ? offset : 0),
		};
		const end = {
			col: horizontal ? range.end.col : start.col,
			row: horizontal ? start.row : range.end.row,
		};
		const first = absolute(start.col, start.row),
			last = absolute(end.col, end.row);
		return {
			formula: `${quoteChartSheetName(range.sheetName)}!${first === last ? first : `${first}:${last}`}`,
			isNumeric: dimension.kind === 'numDim',
			level,
			cellCount: count,
		};
	});
}

const allData = (space: ChartExSpace): ChartExData[] => [
	...space.chartData.data,
	...space.plotArea.series.flatMap((series) => (series.data ? [series.data] : [])),
];

const levelRanges = (space: ChartExSpace, resolve?: ChartExFormulaResolver) =>
	allData(space).flatMap((data) =>
		data.dimensions.flatMap((dimension) => dimensionRanges(dimension, resolve)),
	);

/**
 * Workbook writes for every authored dimension and linked name or title: all cached points, or with
 * `before` (the chart as loaded) only the cells the edit changed.
 */
export function chartExSpaceWorkbookWrites(
	space: ChartExSpace,
	resolve?: ChartExFormulaResolver,
	before?: ChartExSpace,
): ChartWorkbookWrite[] {
	const prior = new Map(
		before ? levelRanges(before, resolve).map((item) => [item.formula, item.level]) : [],
	);
	const writes: ChartWorkbookWrite[] = [];
	for (const { formula, isNumeric, level, cellCount } of levelRanges(space, resolve)) {
		const cells = chartCacheCellChanges(level, prior.get(formula));
		if (cells.length || !before) writes.push({ formula, isNumeric, cells, cellCount });
	}
	const textOf = (tx: ChartExText) => tx.value ?? tx.rich?.text ?? '';
	const texts = (from: ChartExSpace) => [from.title?.tx, ...from.plotArea.series.map((s) => s.tx)];
	const priorTexts = new Map(
		(before ? texts(before) : []).flatMap((tx) =>
			tx?.formula ? [[tx.formula.formula, textOf(tx)] as const] : [],
		),
	);
	for (const text of texts(space))
		if (text?.formula) {
			const r = chartExLocalRange(text.formula.formula, resolve);
			if (r.start.row !== r.end.row || r.start.col !== r.end.col)
				throw new Error('ChartEx text source must be one cell.');
			const value = textOf(text);
			if (before && priorTexts.get(text.formula.formula) === value) continue;
			writes.push({
				formula: `${quoteChartSheetName(r.sheetName)}!${absolute(r.start.col, r.start.row)}`,
				isNumeric: false,
				cells: [{ index: 0, value }],
				cellCount: 1,
			});
		}
	return writes;
}

/** Whether an edit changes source data rather than chart-only presentation. */
export function chartExDataChanged(before: ChartExSpace, after: ChartExSpace): boolean {
	const data = (data: ChartExSpace['chartData']['data'][number]) => ({
		id: data.id,
		dimensions: data.dimensions.map((dimension) => ({
			kind: dimension.kind,
			formula: dimension.formula,
			nameFormula: dimension.nameFormula,
			levels: dimension.levels.map((level) => ({
				pointCount: level.pointCount,
				points: level.points,
				name: level.name,
			})),
		})),
	});
	const shape = (space: ChartExSpace) =>
		JSON.stringify({
			data: space.chartData.data.map(data),
			series: space.plotArea.series.map((series) => ({
				tx: series.tx,
				data: series.data ? data(series.data) : undefined,
				dataId: series.dataId,
			})),
		});
	return shape(before) !== shape(after);
}
