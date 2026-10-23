/**
 * The standard layout of a chart's data in its embedded workbook, the one Word, PowerPoint and
 * Excel write for a new chart: category labels in column A from row 2, series names across row 1
 * from column B, and each series' values in its column under its name. A chart whose formulas
 * follow it can grow and shrink with its data sheet; any other layout is written cell for cell
 * only while its ranges keep their size.
 *
 * @module chart/workbook/layout
 */
import type { ChartSpace } from '../model';
import type { ChartSpaceSeries } from '../model-series';
import { formatCellAddress, parseChartFormulaRange, type ChartFormulaRange } from './cellref';

/** The sheet a new chart's data is written to, Office's default name. */
export const DEFAULT_CHART_SHEET_NAME = 'Sheet1';

/** A sheet name as a formula writes it: quoted when it is not a plain identifier. */
export function quoteChartSheetName(name: string): string {
	return /^[A-Za-z_][A-Za-z0-9_.]*$/u.test(name) ? name : `'${name.replace(/'/gu, "''")}'`;
}

const absolute = (col: number, row: number) =>
	formatCellAddress({ col, row }).replace(/^([A-Z]+)(\d+)$/u, '$$$1$$$2');

function range(sheet: string, col: number, fromRow: number, toRow: number): string {
	const start = absolute(col, fromRow);
	return `${quoteChartSheetName(sheet)}!${fromRow === toRow ? start : `${start}:${absolute(col, toRow)}`}`;
}

/** The formulas of series `index` (0-based) over `categoryCount` categories in the standard layout. */
export interface ChartSeriesFormulas {
	name: string;
	categories: string;
	values: string;
}

/** Formulas of series `index` in the standard layout of `sheet`. */
export function standardChartFormulas(
	sheet: string,
	index: number,
	categoryCount: number,
): ChartSeriesFormulas {
	const last = Math.max(1, categoryCount);
	return {
		name: range(sheet, index + 1, 0, 0),
		categories: range(sheet, 0, 1, last),
		values: range(sheet, index + 1, 1, last),
	};
}

const isXY = (kind: string) => kind === 'scatter' || kind === 'bubble';

function sameRange(formula: string | undefined, expected: string): boolean {
	if (formula === undefined) return false;
	const a = parseChartFormulaRange(formula);
	const b = parseChartFormulaRange(expected) as ChartFormulaRange;
	return (
		a !== undefined &&
		a.sheetName.toLowerCase() === b.sheetName.toLowerCase() &&
		a.start.col === b.start.col &&
		a.start.row === b.start.row &&
		a.end.col === b.end.col &&
		a.end.row === b.end.row
	);
}

/** Every series with the keys its categories and values live under. */
function seriesWithKeys(space: ChartSpace) {
	return space.plotArea.groups.flatMap((group) =>
		group.series.map((series) => ({
			series,
			cat: isXY(group.kind) ? ('xValues' as const) : ('categories' as const),
			val: isXY(group.kind) ? ('yValues' as const) : ('values' as const),
		})),
	);
}

/**
 * The sheet name when every series of `space` (`categoryCount` categories) reads its name,
 * categories and values from the standard layout of one sheet; otherwise `undefined`. A series
 * name given as a literal (`c:v`) is allowed.
 */
export function standardChartLayoutSheet(
	space: ChartSpace,
	categoryCount: number,
): string | undefined {
	const entries = seriesWithKeys(space);
	const first = entries[0]?.series[entries[0].val]?.formula;
	const sheet = first ? parseChartFormulaRange(first)?.sheetName : undefined;
	if (!sheet) return undefined;
	const standard = entries.every(({ series, cat, val }, index) => {
		const expected = standardChartFormulas(sheet, index, categoryCount);
		const name = series.tx?.reference;
		return (
			sameRange(series[val]?.formula, expected.values) &&
			(series[cat] === undefined || sameRange(series[cat]?.formula, expected.categories)) &&
			(name === undefined || sameRange(name.formula, expected.name))
		);
	});
	return standard ? sheet : undefined;
}

function setFormula(
	series: ChartSpaceSeries,
	key: 'categories' | 'values' | 'xValues' | 'yValues',
	formula: string,
) {
	const source = series[key];
	if (!source) return;
	source.formula = formula;
	if (source.kind === 'numLit') source.kind = 'numRef';
	if (source.kind === 'strLit') source.kind = 'strRef';
}

/**
 * Points every series of `space` (mutated) at the standard layout of `sheet` for
 * `categoryCount` categories: names that are references move to row 1, categories and values to
 * their columns.
 */
export function applyStandardChartLayout(
	space: ChartSpace,
	sheet: string,
	categoryCount: number,
): void {
	seriesWithKeys(space).forEach(({ series, cat, val }, index) => {
		const formulas = standardChartFormulas(sheet, index, categoryCount);
		if (series.tx?.reference) series.tx.reference.formula = formulas.name;
		setFormula(series, cat, formulas.categories);
		setFormula(series, val, formulas.values);
	});
}
