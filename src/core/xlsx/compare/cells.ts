// Cell differences of one pair of sheets. Cells are matched by address, as Spreadsheet Compare
// does, so an inserted row shows as changes to every cell below it. The walk visits each stored
// cell of both sheets once, in row-major order (so a capped result is the same every time); it
// never builds a dense grid.
import { MAX_COL, formatRange } from '../address';
import { dataTableAt, dataTableFormula } from '../edit/data-table';
import { effectiveStyleId } from '../layout/cell-view';
import { needsArrayEvaluation } from '../formula/array-context';
import { isSpilledCell } from '../formula/spill';
import { isCellError, type Cell, type CellValue, type Workbook, type Worksheet } from '../model';
import { FormatIndex, paintedStyle } from './format-key';
import { richTextChanged } from './rich-text';
import type { CellDifference, CellSide, CompareCategory, FormatAttribute } from './types';

/** Counts reported differences against `maxDifferences`. */
export class Budget {
	count = 0;
	truncated = false;
	constructor(private readonly max: number) {}

	/** Takes room for one difference; false (and `truncated`) once the cap is reached. */
	take(): boolean {
		if (this.count >= this.max) {
			this.truncated = true;
			return false;
		}
		this.count++;
		return true;
	}

	/** The leading items that fit. */
	admit<T>(items: T[]): T[] {
		const room = Math.max(0, this.max - this.count);
		if (items.length <= room) {
			this.count += items.length;
			return items;
		}
		this.truncated = true;
		this.count += room;
		return items.slice(0, room);
	}
}

export interface CellCompareContext {
	categories: ReadonlySet<CompareCategory>;
	formats: FormatIndex;
	budget: Budget;
}

/** One sheet of the pair, with what the walk has learnt about it so far. */
class Side {
	/**
	 * The last row of an array formula in each column, once one is met. An anchor is the top-left
	 * cell of its range, so the walk records it before any cell of its results.
	 */
	private arrayEnd: Int32Array | undefined;
	private readonly tables: boolean;

	constructor(private readonly sheet: Worksheet) {
		this.tables = !!sheet.dataTables?.length;
	}

	/** Records an array formula anchor met by the walk. */
	visit(cell: Cell | undefined): void {
		const range = cell?.arrayRange;
		if (!range) return;
		this.arrayEnd ??= new Int32Array(MAX_COL + 1).fill(-1);
		for (let col = range.start.col; col <= range.end.col; col++)
			if ((this.arrayEnd[col] ?? -1) < range.end.row) this.arrayEnd[col] = range.end.row;
	}

	/** The formula a cell shows: its own, or the `TABLE()` of the data table it is a result of. */
	formulaOf(cell: Cell | undefined, row: number, col: number): string | undefined {
		if (cell?.formula !== undefined) return cell.formula;
		const table = this.tables ? dataTableAt(this.sheet, row, col) : undefined;
		return table ? dataTableFormula(table) : undefined;
	}

	/** The formula with its array range, so resizing an array formula changes the formula. */
	formulaKey(cell: Cell | undefined, row: number, col: number): string | undefined {
		const formula = this.formulaOf(cell, row, col);
		return formula !== undefined && cell?.arrayRange
			? `${formula}\n${formatRange(cell.arrayRange)}`
			: formula;
	}

	/** The cell holds a typed constant, not a spilled or array formula result. */
	isTyped(cell: Cell | undefined, row: number, col: number): boolean {
		if (!cell || cell.value === null || isSpilledCell(cell)) return false;
		return (this.arrayEnd?.[col] ?? -1) < row;
	}

	/**
	 * The format a cell shows, as the grid resolves it (`effectiveStyleId`): a stored cell's own
	 * (none is style 0, as for a `<c>` without `s`), an absent cell its row's, then its column's.
	 */
	styleIdAt(cell: Cell | undefined, row: number, col: number): number | undefined {
		return cell ? cell.styleId : effectiveStyleId(this.sheet, row, col);
	}

	cellSide(cell: Cell | undefined, row: number, col: number): CellSide {
		const out: CellSide = { value: cell?.value ?? null };
		const formula = this.formulaOf(cell, row, col);
		if (formula !== undefined) out.formula = formula;
		if (cell?.arrayRange) out.arrayRange = cell.arrayRange;
		return out;
	}
}

/** Sorted, distinct keys of up to two maps. */
function sortedKeys(
	a: Map<number, unknown> | undefined,
	b: Map<number, unknown> | undefined,
): Int32Array {
	const keys = new Int32Array((a?.size ?? 0) + (b?.size ?? 0));
	let n = 0;
	if (a) for (const key of a.keys()) keys[n++] = key;
	if (b) for (const key of b.keys()) if (!a?.has(key)) keys[n++] = key;
	return keys.subarray(0, n).sort();
}

const sameValue = (a: CellValue, b: CellValue): boolean =>
	a === b || (isCellError(a) && isCellError(b) && a.error === b.error);

const isLegacy = (cell: Cell | undefined) => !!cell?.legacyFormula && !cell.arrayRange;

/**
 * The same formula switched between dynamic-array and legacy (implicit intersection) semantics
 * where that changes its result: only a formula that needs array evaluation, the test the writer
 * uses to store a formula as a dynamic array (a loaded `A1+1` is legacy and means the same).
 */
function legacyChanged(b: Cell | undefined, a: Cell | undefined): boolean {
	return (
		isLegacy(b) !== isLegacy(a) &&
		b?.formula !== undefined &&
		b.formula === a?.formula &&
		needsArrayEvaluation(b.formula)
	);
}

/** A check of the before (`b`) and after (`a`) cells at one address. */
type AtAddress<T> = (b: Cell | undefined, a: Cell | undefined, row: number, col: number) => T;

/** Differences between two sheets matched by address. */
export function compareSheetCells(
	ctx: CellCompareContext,
	beforeBook: Workbook,
	beforeSheet: Worksheet,
	afterBook: Workbook,
	afterSheet: Worksheet,
): CellDifference[] {
	const before = new Side(beforeSheet);
	const after = new Side(afterSheet);
	const has = (category: CompareCategory) => ctx.categories.has(category);
	const values = has('enteredValue') || has('formula') || has('calculatedValue');
	const formats = has('format');
	const out: CellDifference[] = [];

	/** How the values of the cells at one address differ; undefined when they agree. */
	const valueKind: AtAddress<CellDifference['kind'] | undefined> = (b, a, row, col) => {
		const formula = before.formulaKey(b, row, col);
		if (formula !== after.formulaKey(a, row, col) || legacyChanged(b, a)) return 'formula';
		if (sameValue(b?.value ?? null, a?.value ?? null)) return undefined;
		if (formula === undefined && (before.isTyped(b, row, col) || after.isTyped(a, row, col)))
			return 'enteredValue';
		return 'calculatedValue';
	};

	const valueDifference: AtAddress<CellDifference | undefined> = (b, a, row, col) => {
		const kind = valueKind(b, a, row, col);
		if (!kind || !has(kind)) return undefined;
		const sides = { before: before.cellSide(b, row, col), after: after.cellSide(a, row, col) };
		if (kind === 'formula' && legacyChanged(b, a))
			(isLegacy(b) ? sides.before : sides.after).legacyFormula = true;
		return { kind, address: { row, col }, ...sides };
	};

	const formatDifference: AtAddress<CellDifference | undefined> = (b, a, row, col) => {
		const bStyle = before.styleIdAt(b, row, col);
		const aStyle = after.styleIdAt(a, row, col);
		const changed = ctx.formats.attributesBetween(
			ctx.formats.idOf(beforeBook, bStyle),
			ctx.formats.idOf(afterBook, aStyle),
		);
		// The formats as compared (and painted): a font missing a name or size has style 0's.
		const bFormat = paintedStyle(beforeBook, bStyle);
		const aFormat = paintedStyle(afterBook, aStyle);
		const runs = richTextChanged(b, a, bFormat.font, aFormat.font);
		const attributes: FormatAttribute[] = runs ? [...changed, 'richText'] : changed;
		if (!attributes.length) return undefined;
		return {
			kind: 'format',
			address: { row, col },
			before: { ...before.cellSide(b, row, col), style: bFormat },
			after: { ...after.cellSide(a, row, col), style: aFormat },
			attributes,
		};
	};

	/** Adds a difference while the budget lasts; false once it has run out. */
	const report = (difference: CellDifference | undefined): boolean => {
		if (!difference) return true;
		if (!ctx.budget.take()) return false;
		out.push(difference);
		return true;
	};

	for (const row of sortedKeys(beforeSheet.rows, afterSheet.rows)) {
		const bRow = beforeSheet.rows.get(row);
		const aRow = afterSheet.rows.get(row);
		for (const col of sortedKeys(bRow, aRow)) {
			const b = bRow?.get(col);
			const a = aRow?.get(col);
			before.visit(b);
			after.visit(a);
			if (values && !report(valueDifference(b, a, row, col))) return out;
			if (formats && !report(formatDifference(b, a, row, col))) return out;
		}
	}
	return out;
}
