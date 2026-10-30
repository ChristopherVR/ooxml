import { type CellAddress, type CellRange, normalizeRange, rangeContains } from '../address';
import { getCell } from '../cells';
import type { AutoFilter, AutoFilterColumn, FilterCriteria, Worksheet } from '../model';
import { type EditContext, sheetAt } from './context';
import { compileAutoFilter, setRowFiltered, type FilterEvalOptions } from '../filter-rows';
import { createConditionalFormatEvaluator } from '../layout/cf-evaluator';
import { sortRows } from './sort';

/**
 * The current region around a cell (Ctrl+A / Ctrl+Shift+8): the smallest rectangle bounded by
 * empty rows and columns that contains the cell.
 */
export function currentRegion(sheet: Worksheet, at: CellAddress): CellRange {
	const filled = (row: number, col: number): boolean => {
		const cell = getCell(sheet, row, col);
		return !!cell && cell.value !== null && cell.value !== '';
	};
	let r: CellRange = { start: { ...at }, end: { ...at } };
	for (let grew = true; grew;) {
		grew = false;
		const rowHas = (row: number): boolean => {
			for (let c = Math.max(0, r.start.col - 1); c <= r.end.col + 1; c++)
				if (filled(row, c)) return true;
			return false;
		};
		const colHas = (col: number): boolean => {
			for (let row = Math.max(0, r.start.row - 1); row <= r.end.row + 1; row++)
				if (filled(row, col)) return true;
			return false;
		};
		if (r.start.row > 0 && rowHas(r.start.row - 1)) {
			r = { start: { ...r.start, row: r.start.row - 1 }, end: r.end };
			grew = true;
		}
		if (rowHas(r.end.row + 1)) {
			r = { start: r.start, end: { ...r.end, row: r.end.row + 1 } };
			grew = true;
		}
		if (r.start.col > 0 && colHas(r.start.col - 1)) {
			r = { start: { ...r.start, col: r.start.col - 1 }, end: r.end };
			grew = true;
		}
		if (colHas(r.end.col + 1)) {
			r = { start: r.start, end: { ...r.end, col: r.end.col + 1 } };
			grew = true;
		}
	}
	return r;
}

/** Shows or hides the data rows of an auto-filter according to its column filters. */
export function applyFilter(ctx: EditContext, sheet: Worksheet, filter: AutoFilter): void {
	const { range } = filter;
	const matches = compileAutoFilter(ctx.workbook, sheet, filter, filterOptions(ctx, sheet, filter));
	for (let row = range.start.row + 1; row <= range.end.row; row++)
		setRowFiltered(sheet, row, !matches(row), !!filter.columns?.length);
}

/** Conditional-format output for colour and icon filters, built only when a column needs it. */
function filterOptions(ctx: EditContext, sheet: Worksheet, filter: AutoFilter): FilterEvalOptions {
	const needs = filter.columns?.some((c) => c.color || c.icon);
	const index = ctx.workbook.sheets.indexOf(sheet);
	if (!needs || index < 0 || !sheet.conditionalFormats.length) return {};
	try {
		const cf = createConditionalFormatEvaluator(ctx.workbook, index, (formula, at) =>
			ctx.calc.evaluate(formula, at),
		);
		return { formatAt: (row, col) => cf.at(row, col) };
	} catch {
		return {};
	}
}

/** Turns the auto-filter on for `range` (a single cell expands to its region) or off. */
export function setAutoFilter(ctx: EditContext, s: number, range: CellRange | undefined): void {
	const sheet = sheetAt(ctx.workbook, s);
	let r = range && normalizeRange(range);
	if (r && r.start.row === r.end.row && r.start.col === r.end.col)
		r = currentRegion(sheet, r.start);
	ctx.run(
		r ? 'Filter' : 'Clear filter',
		'view',
		[{ kind: 'sheet', sheet: s }],
		() => {
			const old = sheet.autoFilter;
			if (old) applyFilter(ctx, sheet, { range: old.range });
			if (r) sheet.autoFilter = { range: r };
			else Reflect.deleteProperty(sheet, 'autoFilter');
		},
		{ sheet: s, structural: true },
	);
}

/** Filters one auto-filter column to the given display texts (`''` keeps blanks). */
export function filterColumn(
	ctx: EditContext,
	s: number,
	col: number,
	values: string[] | undefined,
): void {
	setColumnCriteria(
		ctx,
		s,
		col,
		values && {
			values: values.filter((v) => v !== ''),
			...(values.includes('') ? { blank: true as const } : {}),
		},
	);
}

/** Sets (or with `undefined` clears) any criteria of one auto-filter column and reapplies. */
export function setColumnCriteria(
	ctx: EditContext,
	s: number,
	col: number,
	criteria: FilterCriteria | undefined,
): void {
	const sheet = sheetAt(ctx.workbook, s);
	const filter = sheet.autoFilter;
	if (!filter || col < filter.range.start.col || col > filter.range.end.col)
		throw new Error('The column is not inside an auto-filter range.');
	const offset = col - filter.range.start.col;
	ctx.run(
		criteria ? 'Filter column' : 'Clear column filter',
		'view',
		[{ kind: 'sheet', sheet: s }],
		() => {
			const old = filter.columns?.find((fc) => fc.offset === offset);
			const columns = (filter.columns ?? []).filter((fc) => fc.offset !== offset);
			if (criteria) {
				const entry: AutoFilterColumn = { offset, ...structuredClone(criteria) };
				if (old?.hiddenButton) entry.hiddenButton = true;
				columns.push(entry);
			}
			columns.sort((a, b) => a.offset - b.offset);
			const next: AutoFilter = { ...filter };
			if (columns.length) next.columns = columns;
			else delete next.columns;
			sheet.autoFilter = next;
			applyFilter(ctx, sheet, next);
		},
		{ sheet: s, structural: true },
	);
}

/** Re-evaluates every column criterion against the current cell values (Data > Reapply). */
export function reapplyFilter(ctx: EditContext, s: number): void {
	const sheet = sheetAt(ctx.workbook, s);
	const filter = sheet.autoFilter;
	if (!filter) return;
	ctx.run('Reapply', 'view', [{ kind: 'sheet', sheet: s }], () => applyFilter(ctx, sheet, filter), {
		sheet: s,
		structural: true,
	});
}

/** Sorts the table or auto-filter range containing `col` by that column, keeping its header. */
export function sortByColumn(ctx: EditContext, s: number, col: number, descending: boolean): void {
	const sheet = sheetAt(ctx.workbook, s);
	const table = sheet.tables.find((t) => col >= t.range.start.col && col <= t.range.end.col);
	const filter = sheet.autoFilter;
	const target = table
		? { range: table.range, header: table.headerRow, totals: table.totalsRow }
		: filter && rangeContains(filter.range, { row: filter.range.start.row, col })
			? { range: filter.range, header: true, totals: false }
			: undefined;
	if (!target) throw new Error('The column is not inside a table or an auto-filter range.');
	const range = target.totals
		? { start: target.range.start, end: { ...target.range.end, row: target.range.end.row - 1 } }
		: target.range;
	ctx.run(
		descending ? 'Sort Z to A' : 'Sort A to Z',
		'cells',
		[{ kind: 'sheet', sheet: s }],
		() => sortRows(sheet, range, [{ col, descending }], target.header),
		{ sheet: s, ranges: [range] },
	);
}
