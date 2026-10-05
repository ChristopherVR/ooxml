// Scalar operators and Excel's array broadcasting.
import { compareScalars, toNumber, toText } from './coerce.js';
import { ERR, ErrorSignal, isError, Matrix, type Scalar } from './values.js';

export type ScalarOp = (a: Scalar, b: Scalar) => Scalar;

const signal = (fn: () => Scalar): Scalar => {
	try {
		return fn();
	} catch (e) {
		if (e instanceof ErrorSignal) return e.value;
		throw e;
	}
};

/** `(-8)^(1/3)` is -2 in Excel: odd roots of negative numbers are allowed (POWER shares it). */
export function power(base: number, exponent: number): number {
	if (base === 0 && exponent === 0) throw new ErrorSignal(ERR.NUM);
	if (base === 0 && exponent < 0) throw new ErrorSignal(ERR.DIV0);
	if (base < 0 && !Number.isInteger(exponent)) {
		const inverse = 1 / exponent;
		const odd = Math.round(inverse);
		if (Math.abs(inverse - odd) < 1e-10 && Math.abs(odd) % 2 === 1) {
			return -Math.pow(-base, exponent);
		}
		throw new ErrorSignal(ERR.NUM);
	}
	return Math.pow(base, exponent);
}

/** Excel sets a final addition or subtraction that cancels to within 15 digits to exactly 0. */
export function snapToZero(result: number, a: number, b: number): number {
	if (result !== 0 && Math.abs(result) <= Math.max(Math.abs(a), Math.abs(b)) * 1e-15) return 0;
	return result;
}

function arithmetic(op: string, x: Scalar, y: Scalar, snap: boolean): Scalar {
	if (isError(x)) return x;
	if (isError(y)) return y;
	return signal(() => {
		const a = toNumber(x);
		const b = toNumber(y);
		let r: number;
		switch (op) {
			case '+':
				r = a + b;
				if (snap) r = snapToZero(r, a, b);
				break;
			case '-':
				r = a - b;
				if (snap) r = snapToZero(r, a, b);
				break;
			case '*':
				r = a * b;
				break;
			case '/':
				if (b === 0) return ERR.DIV0;
				r = a / b;
				break;
			default:
				r = power(a, b);
		}
		return Number.isFinite(r) ? (r === 0 ? 0 : r) : ERR.NUM;
	});
}

function comparison(op: string, x: Scalar, y: Scalar): Scalar {
	if (isError(x)) return x;
	if (isError(y)) return y;
	const c = compareScalars(x, y);
	switch (op) {
		case '=':
			return c === 0;
		case '<>':
			return c !== 0;
		case '<':
			return c < 0;
		case '>':
			return c > 0;
		case '<=':
			return c <= 0;
		default:
			return c >= 0;
	}
}

function concat(x: Scalar, y: Scalar): Scalar {
	if (isError(x)) return x;
	if (isError(y)) return y;
	return toText(x) + toText(y);
}

/** Applies a binary operator to two scalars. */
export function scalarBinary(op: string, x: Scalar, y: Scalar, snap = false): Scalar {
	if (op === '&') return concat(x, y);
	if (op === '+' || op === '-' || op === '*' || op === '/' || op === '^') {
		return arithmetic(op, x, y, snap);
	}
	return comparison(op, x, y);
}

/** Unary minus and percent. */
export function scalarUnary(op: '-' | '%', x: Scalar): Scalar {
	if (isError(x)) return x;
	return signal(() => {
		const n = toNumber(x);
		const r = op === '-' ? -n : n / 100;
		return r === 0 ? 0 : r;
	});
}

/** The element of a broadcast operand at a position (size-1 dimensions repeat; outside is #N/A). */
export function pick(m: Matrix, row: number, col: number): Scalar {
	const r = m.rows === 1 ? 0 : row;
	const c = m.cols === 1 ? 0 : col;
	if (r >= m.rows || c >= m.cols) return ERR.NA;
	return m.get(r, c);
}

/**
 * The elementwise result over broadcast operands, `at(r, c)` giving one element. When an operand
 * is padded (a whole-column read), only the block the operands store is computed and the rest
 * becomes padding holding the value of one tail element, as long as every operand is uniform
 * there; otherwise the full result is built (within the array size limits).
 */
export function broadcastShape(operands: Matrix[], at: (r: number, c: number) => Scalar): Matrix {
	let rows = 1;
	let cols = 1;
	for (const m of operands) {
		rows = Math.max(rows, m.rows);
		cols = Math.max(cols, m.cols);
	}
	if (!operands.some((m) => m.isPadded)) return Matrix.build(rows, cols, at);
	let blockRows = 1;
	let blockCols = 1;
	for (const m of operands) {
		blockRows = Math.max(blockRows, m.isPadded ? m.blockRows : m.rows);
		blockCols = Math.max(blockCols, m.isPadded ? m.blockCols : m.cols);
	}
	blockRows = Math.min(blockRows, rows);
	blockCols = Math.min(blockCols, cols);
	// A single row repeated down the row padding (or a column across the column padding) varies.
	const varies = operands.some(
		(m) =>
			(blockRows < rows && m.rows === 1 && m.cols > 1) ||
			(blockCols < cols && m.cols === 1 && m.rows > 1),
	);
	if (varies) return Matrix.build(rows, cols, at);
	return Matrix.padded(Matrix.build(blockRows, blockCols, at), rows, cols, at(rows - 1, cols - 1));
}

/** Broadcasts a binary scalar function over matrices like Excel's array arithmetic. */
export function broadcast2(a: Scalar | Matrix, b: Scalar | Matrix, fn: ScalarOp): Scalar | Matrix {
	if (!(a instanceof Matrix) && !(b instanceof Matrix)) return fn(a, b);
	const ma = a instanceof Matrix ? a : new Matrix([[a]]);
	const mb = b instanceof Matrix ? b : new Matrix([[b]]);
	return broadcastShape([ma, mb], (r, c) => fn(pick(ma, r, c), pick(mb, r, c)));
}

/** Maps a unary scalar function over a matrix. */
export function broadcast1(a: Scalar | Matrix, fn: (x: Scalar) => Scalar): Scalar | Matrix {
	return a instanceof Matrix ? a.mapValues(fn) : fn(a);
}
