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

/** A dense two-dimensional array of values (array constants, array results, dereferenced ranges). */
export class Matrix {
	constructor(readonly data: Scalar[][]) {}

	get rows(): number {
		return this.data.length;
	}

	get cols(): number {
		return this.data[0]?.length ?? 0;
	}

	get(row: number, col: number): Scalar {
		return this.data[row]?.[col] ?? null;
	}

	static build(rows: number, cols: number, fill: (row: number, col: number) => Scalar): Matrix {
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

	/** Every element in row-major order. */
	flat(): Scalar[] {
		return this.data.flat();
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
