// Map ChartEx caches to exact local workbook vectors, including multi-column hierarchies.
import type { ChartExDimension, ChartExSpace } from './model';
import { chartExLevelValues } from '../edit/chartex-sheet';
import { formatCellAddress, parseChartFormulaRange } from '../workbook/cellref';
import { quoteChartSheetName } from '../workbook/layout';
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

function dimensionWrites(
	dimension: ChartExDimension,
	resolve?: ChartExFormulaResolver,
): ChartWorkbookWrite[] {
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
		const values = chartExLevelValues(level);
		if (values.length !== count)
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
			values,
		};
	});
}
/** Atomic workbook writes implied by every authored dimension and linked name/title. */
export function chartExSpaceWorkbookWrites(
	space: ChartExSpace,
	resolve?: ChartExFormulaResolver,
): ChartWorkbookWrite[] {
	const writes: ChartWorkbookWrite[] = [];
	for (const data of [
		...space.chartData.data,
		...space.plotArea.series.flatMap((series) => (series.data ? [series.data] : [])),
	])
		for (const dimension of data.dimensions) writes.push(...dimensionWrites(dimension, resolve));
	for (const text of [space.title?.tx, ...space.plotArea.series.map((series) => series.tx)])
		if (text?.formula) {
			const r = chartExLocalRange(text.formula.formula, resolve);
			if (r.start.row !== r.end.row || r.start.col !== r.end.col)
				throw new Error('ChartEx text source must be one cell.');
			writes.push({
				formula: `${quoteChartSheetName(r.sheetName)}!${absolute(r.start.col, r.start.row)}`,
				isNumeric: false,
				values: [text.value ?? text.rich?.text ?? ''],
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
