import type { CellRange } from '../address.js';
import type { CellError, ErrorCode } from '../model.js';
import { isCellError } from '../model.js';
import type { FormulaAst } from './ast.js';

/** A single value as formulas see it; `null` is an empty cell (or an omitted argument). */
export type Scalar = number | string | boolean | CellError | null;

const ERROR_CACHE = new Map<ErrorCode, CellError>();

/** The shared, frozen error value for `code`. */
export function err(code: ErrorCode): CellError {
	let value = ERROR_CACHE.get(code);
	if (!value) {
		value = Object.freeze({ error: code });
		ERROR_CACHE.set(code, value);
	}
	return value;
}

export const ERR = {
	NULL: err('#NULL!'),
	DIV0: err('#DIV/0!'),
	VALUE: err('#VALUE!'),
	REF: err('#REF!'),
	NAME: err('#NAME?'),
	NUM: err('#NUM!'),
	NA: err('#N/A'),
	SPILL: err('#SPILL!'),
	CALC: err('#CALC!'),
} as const;

export const isError = isCellError;

/** Excel's longest array side: SEQUENCE(1048577) and SEQUENCE(1,1048577) are `#VALUE!`. */
export const MAX_ARRAY_SIDE = 1_048_576;
/**
 * The most cells one array may hold, so no formula can exhaust memory (Excel gives up somewhere
 * past 40 million with "out of resources"; this engine stops at 16 million with `#NUM!`).
 */
export const MAX_ARRAY_CELLS = 16_777_216;

/** `#VALUE!` past the longest side, `#NUM!` past the cell budget. */
export function checkArraySize(rows: number, cols: number): void {
	if (rows > MAX_ARRAY_SIDE || cols > MAX_ARRAY_SIDE) fail(ERR.VALUE);
	if (rows * cols > MAX_ARRAY_CELLS) fail(ERR.NUM);
}

/** The part of a padded matrix outside its stored block: every such cell holds `fill`. */
interface Padding {
	rows: number;
	cols: number;
	fill: Scalar;
}

/**
 * A two-dimensional array of values (array constants, array results, dereferenced ranges). A
 * padded matrix stores only a top-left block and reports a larger size whose remaining cells all
 * hold one value: whole-column references read that way, so `ROWS(A:A*1)` is 1048576 and
 * `SUMPRODUCT(--(A:A=""))` counts every blank row without materialising a million cells.
 */
export class Matrix {
	private readonly dense: Scalar[][];
	private readonly padding: Padding | undefined;
	/** The full rows of a padded matrix, built once when `data` is first asked for. */
	private materialized: Scalar[][] | undefined;

	constructor(data: Scalar[][], padding?: Padding) {
		this.dense = data;
		const rows = data.length;
		const cols = data[0]?.length ?? 0;
		this.padding =
			padding && (padding.rows > rows || padding.cols > cols) ? { ...padding } : undefined;
	}

	/** A padded matrix (or a plain one when `rows` x `cols` is no larger than the block). */
	static padded(block: Matrix, rows: number, cols: number, fill: Scalar): Matrix {
		return new Matrix(block.dense, { rows, cols, fill });
	}

	/** Every row, materialising a padded matrix (bounded by `MAX_ARRAY_CELLS`). */
	get data(): Scalar[][] {
		const pad = this.padding;
		if (!pad) return this.dense;
		if (this.materialized) return this.materialized;
		checkArraySize(pad.rows, pad.cols);
		const out: Scalar[][] = [];
		for (let r = 0; r < pad.rows; r++) {
			const line: Scalar[] = [];
			const stored = this.dense[r];
			for (let c = 0; c < pad.cols; c++)
				line.push(stored && c < stored.length ? (stored[c] as Scalar) : pad.fill);
			out.push(line);
		}
		this.materialized = out;
		return out;
	}

	get rows(): number {
		return this.padding ? this.padding.rows : this.dense.length;
	}

	get cols(): number {
		return this.padding ? this.padding.cols : (this.dense[0]?.length ?? 0);
	}

	/** Whether some cells are implied by the padding rather than stored. */
	get isPadded(): boolean {
		return this.padding !== undefined;
	}

	/** Rows and columns of the stored block. */
	get blockRows(): number {
		return this.dense.length;
	}

	get blockCols(): number {
		return this.dense[0]?.length ?? 0;
	}

	/** The value of every cell outside the stored block. */
	get fill(): Scalar {
		return this.padding?.fill ?? null;
	}

	get(row: number, col: number): Scalar {
		const stored = this.dense[row];
		if (stored && col < stored.length) return stored[col] ?? null;
		const pad = this.padding;
		return pad && row < pad.rows && col < pad.cols ? pad.fill : null;
	}

	/** Only the stored block of a padded matrix (what a spilling formula shows). */
	block(): Matrix {
		return this.padding ? new Matrix(this.dense) : this;
	}

	static build(rows: number, cols: number, fill: (row: number, col: number) => Scalar): Matrix {
		checkArraySize(rows, cols);
		const data: Scalar[][] = [];
		for (let r = 0; r < rows; r++) {
			const line: Scalar[] = [];
			for (let c = 0; c < cols; c++) line.push(fill(r, c));
			data.push(line);
		}
		return new Matrix(data);
	}

	map(fn: (value: Scalar, row: number, col: number) => Scalar): Matrix {
		return new Matrix(this.data.map((line, r) => line.map((value, c) => fn(value, r, c))));
	}

	/** Maps every value; a padded matrix maps its block and its fill once. */
	mapValues(fn: (value: Scalar) => Scalar): Matrix {
		const block = this.dense.map((line) => line.map((value) => fn(value)));
		const pad = this.padding;
		return new Matrix(block, pad ? { ...pad, fill: fn(pad.fill) } : undefined);
	}

	/** Visits every value in row-major order without materialising a padded matrix. */
	forEachValue(visit: (value: Scalar) => void): void {
		const pad = this.padding;
		if (!pad) {
			for (const line of this.dense) for (const v of line) visit(v);
			return;
		}
		for (let r = 0; r < pad.rows; r++) {
			const stored = this.dense[r];
			const n = stored ? stored.length : 0;
			for (let c = 0; c < n; c++) visit((stored as Scalar[])[c] ?? null);
			for (let c = n; c < pad.cols; c++) visit(pad.fill);
		}
	}

	/** Every element in row-major order. */
	flat(): Scalar[] {
		if (!this.padding) return this.dense.flat();
		const out: Scalar[] = [];
		checkArraySize(this.rows, this.cols);
		this.forEachValue((v) => out.push(v));
		return out;
	}
}

/** One rectangular area of a reference. */
export interface Area {
	sheet: number;
	range: CellRange;
}

/** A reference to one or more areas (a union when there are several). */
export class RefValue {
	constructor(readonly areas: readonly Area[]) {}

	get first(): Area | undefined {
		return this.areas[0];
	}

	/** Whether this is a single cell. */
	isCell(): boolean {
		const area = this.areas[0];
		return (
			this.areas.length === 1 &&
			area !== undefined &&
			area.range.start.row === area.range.end.row &&
			area.range.start.col === area.range.end.col
		);
	}
}

/** Lexical variables of LET and LAMBDA, keyed by upper-case name. */
export type Scope = ReadonlyMap<string, Value>;

/** A LAMBDA closure. */
export class LambdaValue {
	constructor(
		readonly params: readonly string[],
		readonly body: FormulaAst,
		readonly scope: Scope | undefined,
	) {}
}

export type Value = Scalar | Matrix | RefValue | LambdaValue;

/** Thrown inside function implementations to return an error value. */
export class ErrorSignal {
	constructor(readonly value: CellError) {}
}

/** Throws the error value `value` (or the error with `code`). */
export function fail(value: CellError | ErrorCode): never {
	throw new ErrorSignal(typeof value === 'string' ? err(value) : value);
}

export const isScalar = (value: Value): value is Scalar =>
	!(value instanceof Matrix || value instanceof RefValue || value instanceof LambdaValue);
