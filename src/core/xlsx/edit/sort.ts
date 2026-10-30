import { type CellRange, formatRange, normalizeRange, rangesIntersect } from '../address';
import { deleteCell, forEachCellInRange, putCell, usedRange } from '../cells';
import { colorFilterCriteria } from '../filter-menu';
import { createConditionalFormatEvaluator } from '../layout/cf-evaluator';
import type { Cell, SortCondition, SortState, Worksheet } from '../model';
import { type EditContext, sheetAt } from './context';
import { isSpilledCell, translateFormula } from './deps';
import { cellRange } from './range-math';
import { compareProbes, probeCell, type Probe, type SortKey, type SortOptions } from './sort-keys';

export { compareCellValues, type SortKey, type SortOptions } from './sort-keys';

const columnHidden = (sheet: Worksheet, col: number): boolean =>
	sheet.columns.some((c) => c.hidden && col >= c.min && col <= c.max);

interface Unit {
	/** Row (top to bottom) or column (left to right) index. */
	pos: number;
	/** The unit's cells keyed by the other axis. */
	cells: Map<number, Cell>;
	probes: Probe[];
}

/**
 * Sorts the visible rows of `range` in place (hidden rows stay put), or its columns when
 * `options.leftToRight`; `hasHeader` keeps the first row (or column) out of the sort.
 */
export function sortRows(
	sheet: Worksheet,
	range: CellRange,
	keys: SortKey[],
	hasHeader: boolean,
	options: SortOptions = {},
): void {
	const across = !!options.leftToRight;
	const r = normalizeRange(range);
	const used = usedRange(sheet);
	const [lo, hi, otherLo, otherHi] = across
		? [r.start.col, r.end.col, r.start.row, r.end.row]
		: [r.start.row, r.end.row, r.start.col, r.end.col];
	const last = Math.min(hi, (across ? used?.end.col : used?.end.row) ?? lo);
	const first = hasHeader ? lo + 1 : lo;
	if (last <= first) return;
	const dataArea = across
		? cellRange(otherLo, first, otherHi, last)
		: cellRange(first, otherLo, last, otherHi);
	if (sheet.merges.some((m) => rangesIntersect(m, dataArea)))
		throw new Error('Cannot sort a range that contains merged cells.');
	let spilled = false;
	forEachCellInRange(sheet, dataArea, (cell) => {
		spilled ||= isSpilledCell(cell);
	});
	if (spilled) throw new Error('Cannot sort a range that contains part of a dynamic array.');
	const sortKeys = keys.length ? keys : [{ col: otherLo }];
	const at = (pos: number, other: number): [number, number] =>
		across ? [other, pos] : [pos, other];
	const units: Unit[] = [];
	for (let pos = first; pos <= last; pos++) {
		if (across ? columnHidden(sheet, pos) : sheet.rowInfo.get(pos)?.hidden) continue;
		const cells = new Map<number, Cell>();
		if (across) {
			for (let row = otherLo; row <= otherHi; row++) {
				const cell = sheet.rows.get(row)?.get(pos);
				if (cell) cells.set(row, cell);
			}
		} else {
			for (const [col, cell] of sheet.rows.get(pos) ?? [])
				if (col >= otherLo && col <= otherHi) cells.set(col, cell);
		}
		const probes = sortKeys.map((key) => {
			const [row, col] = at(pos, key.col);
			return probeCell(key, cells.get(key.col), row, col, options);
		});
		units.push({ pos, cells, probes });
	}
	const sorted = [...units].sort((a, b) => {
		for (let i = 0; i < sortKeys.length; i++) {
			const result = compareProbes(sortKeys[i]!, a.probes[i]!, b.probes[i]!, options);
			if (result) return result;
		}
		return 0;
	});
	for (const { pos, cells } of units)
		for (const other of cells.keys()) deleteCell(sheet, ...at(pos, other));
	moveAnnotations(
		sheet,
		dataArea,
		new Map(sorted.map((source, i) => [source.pos, units[i]!.pos])),
		across,
	);
	sorted.forEach((source, i) => {
		const pos = units[i]!.pos;
		const delta = pos - source.pos;
		for (const [other, cell] of source.cells) {
			if (cell.formula !== undefined && delta) {
				try {
					cell.formula = across
						? translateFormula(cell.formula, 0, delta)
						: translateFormula(cell.formula, delta, 0);
				} catch {
					// Keep a formula the parser cannot read as it was.
				}
			}
			putCell(sheet, ...at(pos, other), cell);
		}
	});
}

/**
 * Comments and single-row (or, left to right, single-column) hyperlinks inside the sorted area
 * travel with their rows or columns (row heights and column widths do not, as in Excel).
 */
function moveAnnotations(
	sheet: Worksheet,
	area: CellRange,
	map: Map<number, number>,
	across: boolean,
): void {
	const inOther = (n: number): boolean =>
		across ? n >= area.start.row && n <= area.end.row : n >= area.start.col && n <= area.end.col;
	for (const comment of sheet.comments) {
		const { row, col } = comment.address;
		const next = map.get(across ? col : row);
		if (next !== undefined && inOther(across ? row : col))
			comment.address = across ? { row, col: next } : { row: next, col };
	}
	for (const link of sheet.hyperlinks) {
		const { start, end } = link.range;
		const pos = across ? start.col : start.row;
		const next = map.get(pos);
		if (next === undefined || pos !== (across ? end.col : end.row)) continue;
		if (!inOther(across ? start.row : start.col) || !inOther(across ? end.row : end.col)) continue;
		link.range = across
			? { start: { row: start.row, col: next }, end: { row: end.row, col: next } }
			: { start: { row: next, col: start.col }, end: { row: next, col: end.col } };
	}
}

/** The saved `sortState` equivalent of the levels just applied (the header is not part of it). */
function sortStateOf(
	range: CellRange,
	keys: SortKey[],
	hasHeader: boolean,
	options: SortOptions,
): SortState | undefined {
	const across = !!options.leftToRight;
	const data = across
		? cellRange(
				range.start.row,
				range.start.col + (hasHeader ? 1 : 0),
				range.end.row,
				range.end.col,
			)
		: cellRange(
				range.start.row + (hasHeader ? 1 : 0),
				range.start.col,
				range.end.row,
				range.end.col,
			);
	if (!keys.length) return undefined;
	const conditions: SortCondition[] = keys.map((key) => {
		const ref = across
			? cellRange(key.col, data.start.col, key.col, data.end.col)
			: cellRange(data.start.row, key.col, data.end.row, key.col);
		const condition: SortCondition = { ref };
		if (key.descending) condition.descending = true;
		if (key.sortOn && key.sortOn !== 'value') condition.sortBy = key.sortOn;
		if (key.customList) condition.customList = key.customList.join(',');
		if (key.sortOn === 'cellColor' || key.sortOn === 'fontColor') {
			const style = colorFilterCriteria(key.color, key.sortOn === 'cellColor').color?.style;
			if (style) condition.style = style;
		}
		if (key.sortOn === 'icon' && key.icon) {
			condition.iconSet = key.icon.set;
			if (key.icon.index !== undefined) condition.iconId = key.icon.index;
		}
		return condition;
	});
	return {
		range: data,
		...(across ? { columnSort: true } : {}),
		...(options.caseSensitive ? { caseSensitive: true } : {}),
		conditions,
	};
}

/** Saves the sort with the sheet's auto-filter when it spans the range, else with the sheet. */
function recordSortState(sheet: Worksheet, range: CellRange, state: SortState | undefined): void {
	const filter = sheet.autoFilter;
	const table = sheet.tables.find((t) => rangesIntersect(t.range, range));
	if (table) {
		if (state) table.sortState = state;
		else Reflect.deleteProperty(table, 'sortState');
		return;
	}
	if (filter && formatRange(filter.range) === formatRange(range)) {
		if (state) filter.sortState = state;
		else Reflect.deleteProperty(filter, 'sortState');
	} else if (state) sheet.sortState = state;
	else Reflect.deleteProperty(sheet, 'sortState');
}

/** Adds the workbook and, for colour and icon levels, the conditional-format output. */
export function resolveSortOptions(
	ctx: EditContext,
	s: number,
	keys: SortKey[],
	options: SortOptions,
): SortOptions {
	const sheet = sheetAt(ctx.workbook, s);
	const resolved: SortOptions = { ...options, workbook: ctx.workbook };
	const needsFormats = keys.some((k) => k.sortOn && k.sortOn !== 'value');
	if (needsFormats && sheet.conditionalFormats.length) {
		try {
			const cf = createConditionalFormatEvaluator(ctx.workbook, s, (formula, at) =>
				ctx.calc.evaluate(formula, at),
			);
			resolved.formatAt = (row, col) => cf.at(row, col);
		} catch {
			// Colour and icon levels then see the static formatting only.
		}
	}
	return resolved;
}

/**
 * Sorts the rows of `range` by `keys` (absolute column indices), stable, header kept in place.
 * With `options.leftToRight` the columns are reordered by the values of the keys' rows. The
 * levels are saved as the sheet's `sortState`.
 */
export function sortRange(
	ctx: EditContext,
	s: number,
	range: CellRange,
	keys: SortKey[],
	hasHeader: boolean,
	options: SortOptions = {},
): void {
	const sheet = sheetAt(ctx.workbook, s);
	const r = normalizeRange(range);
	const resolved = resolveSortOptions(ctx, s, keys, options);
	ctx.run(
		'Sort',
		'cells',
		[{ kind: 'sheet', sheet: s }],
		() => {
			sortRows(sheet, r, keys, hasHeader, resolved);
			recordSortState(sheet, r, sortStateOf(r, keys, hasHeader, options));
		},
		{ sheet: s, ranges: [r] },
	);
}
