// Formula shapes: a formula's text with every digit run blanked. Copies of one formula filled down
// a column (`=B4-C4`, `=B5-C5`, ...) share a shape and differ only in the row numbers of their
// references, so one parsed copy (a template) gives every other copy's tree by replacing those
// rows. Any other digit (a number, a sheet or function name, a string) must match the template
// exactly, and a row out of range or written with a leading zero is refused, so an instantiated
// tree is always the one `parseFormula` would build.
import { MAX_ROW } from '../address';
import { FormulaError, type FormulaAst, type RefSpec } from './ast';
import { parseTokens } from './parser';
import { type Token, tokenize } from './tokenizer';

/** A formula's digit runs (`starts[i]` to `ends[i]`) and a hash of everything else. */
export interface Shape {
	hash: number;
	starts: number[];
	ends: number[];
}

/** Where a reference's rows come from: the indexes of the digit runs holding them. */
interface RowSlots {
	start: number;
	/** Absent for a single cell, whose end corner is its start. */
	end?: number;
}

export interface Template {
	formula: string;
	shape: Shape;
	ast: FormulaAst;
	/** Runs that hold reference rows (all others must match the template exactly). */
	slotted: boolean[];
	refs: Map<RefSpec, RowSlots>;
}

const isDigit = (code: number): boolean => code >= 48 && code <= 57;

export function formulaShape(formula: string): Shape {
	const starts: number[] = [];
	const ends: number[] = [];
	let hash = 0x811c9dc5;
	for (let i = 0; i < formula.length; i++) {
		let code = formula.charCodeAt(i);
		if (isDigit(code)) {
			let j = i + 1;
			while (j < formula.length && isDigit(formula.charCodeAt(j))) j++;
			starts.push(i);
			ends.push(j);
			i = j - 1;
			code = 0;
		}
		hash = Math.imul(hash ^ code, 0x01000193);
	}
	return { hash: hash >>> 0, starts, ends };
}

function sameText(a: string, from: number, b: string, at: number, length: number): boolean {
	for (let k = 0; k < length; k++)
		if (a.charCodeAt(from + k) !== b.charCodeAt(at + k)) return false;
	return true;
}

/** Whether two formulas have the same text apart from their digit runs. */
function sameShape(a: string, as: Shape, b: string, bs: Shape): boolean {
	const runs = as.starts.length;
	if (runs !== bs.starts.length) return false;
	let fromA = 0;
	let fromB = 0;
	for (let i = 0; i <= runs; i++) {
		const toA = i < runs ? (as.starts[i] as number) : a.length;
		const toB = i < runs ? (bs.starts[i] as number) : b.length;
		if (toA - fromA !== toB - fromB || !sameText(a, fromA, b, fromB, toA - fromA)) return false;
		fromA = as.ends[i] ?? 0;
		fromB = bs.ends[i] ?? 0;
	}
	return true;
}

/** Parses a formula and records which of its digit runs are reference rows. */
export function makeTemplate(formula: string, shape: Shape): Template | FormulaError {
	let tokens: Token[];
	let ast: FormulaAst;
	try {
		tokens = tokenize(formula);
		ast = parseTokens(tokens);
	} catch (e) {
		// A stack overflow while parsing is reported like a too-deeply nested formula.
		if (e instanceof RangeError) return new FormulaError(e.message, -1, '#VALUE!');
		if (e instanceof FormulaError) return e;
		throw e;
	}
	const { starts } = shape;
	const slotted = starts.map(() => false);
	const refs = new Map<RefSpec, RowSlots>();
	let run = 0;
	for (const token of tokens) {
		const end = token.start + token.text.length;
		const first = run;
		while (run < starts.length && (starts[run] as number) < end) run++;
		const ref = token.kind === 'ref' ? token.ref : undefined;
		if (!ref || ref.kind === 'cols') continue;
		// The rows are the token's last runs (a sheet prefix may hold digits before them).
		const count = ref.kind === 'cell' ? 1 : 2;
		if (run - first < count) continue;
		const at = run - count;
		const indexes = count === 1 ? [at] : [at, at + 1];
		if (indexes.some((i) => formula.charCodeAt(starts[i] as number) === 48)) continue;
		for (const i of indexes) slotted[i] = true;
		refs.set(ref, count === 1 ? { start: at } : { start: at, end: at + 1 });
	}
	return { formula, shape, ast, slotted, refs };
}

/** The reference rows of a formula of the template's shape, or undefined when it cannot reuse it. */
function rowsFor(t: Template, formula: string, shape: Shape): number[] | undefined {
	if (!sameShape(t.formula, t.shape, formula, shape)) return undefined;
	const rows: number[] = [];
	for (let i = 0; i < shape.starts.length; i++) {
		const start = shape.starts[i] as number;
		const length = (shape.ends[i] as number) - start;
		if (!t.slotted[i]) {
			const own = t.shape.starts[i] as number;
			if ((t.shape.ends[i] as number) - own !== length) return undefined;
			if (!sameText(t.formula, own, formula, start, length)) return undefined;
			rows.push(-1);
			continue;
		}
		if (length > 7 || formula.charCodeAt(start) === 48) return undefined;
		let row = 0;
		for (let k = start; k < start + length; k++) row = row * 10 + formula.charCodeAt(k) - 48;
		if (row < 1 || row > MAX_ROW + 1) return undefined;
		rows.push(row - 1);
	}
	return rows;
}

function shiftSpec(spec: RefSpec, slots: RowSlots, rows: readonly number[]): RefSpec {
	const startRow = rows[slots.start] as number;
	const endRow = slots.end === undefined ? startRow : (rows[slots.end] as number);
	if (startRow === spec.start.row && endRow === spec.end.row) return spec;
	const { start: from, end: to } = spec;
	const start = { row: startRow, col: from.col, rowAbs: from.rowAbs, colAbs: from.colAbs };
	if (slots.end === undefined) return { kind: spec.kind, start, end: start };
	const end = { row: endRow, col: to.col, rowAbs: to.rowAbs, colAbs: to.colAbs };
	return { kind: spec.kind, start, end };
}

/** A copy of the template's tree with its reference rows replaced; unchanged parts are shared. */
function copyTree(node: FormulaAst, t: Template, rows: readonly number[]): FormulaAst {
	switch (node.type) {
		case 'ref': {
			const slots = node.ref ? t.refs.get(node.ref) : undefined;
			if (!slots || !node.ref) return node;
			const ref = shiftSpec(node.ref, slots, rows);
			if (ref === node.ref) return node;
			const copy: FormulaAst = { type: 'ref' };
			if (node.prefix) copy.prefix = node.prefix;
			copy.ref = ref;
			if (node.spill) copy.spill = true;
			return copy;
		}
		case 'unary':
		case 'percent': {
			const operand = copyTree(node.operand, t, rows);
			if (operand === node.operand) return node;
			return node.type === 'unary'
				? { type: 'unary', op: node.op, operand }
				: { type: 'percent', operand };
		}
		case 'binary': {
			const left = copyTree(node.left, t, rows);
			const right = copyTree(node.right, t, rows);
			if (left === node.left && right === node.right) return node;
			return { type: 'binary', op: node.op, left, right };
		}
		case 'call': {
			const args = copyAll(node.args, t, rows);
			if (args === node.args) return node;
			return { type: 'call', name: node.name, rawName: node.rawName, args };
		}
		case 'invoke': {
			const callee = copyTree(node.callee, t, rows);
			const args = copyAll(node.args, t, rows);
			if (callee === node.callee && args === node.args) return node;
			return { type: 'invoke', callee, args };
		}
		default:
			return node;
	}
}

function copyAll(nodes: FormulaAst[], t: Template, rows: readonly number[]): FormulaAst[] {
	let out: FormulaAst[] | undefined;
	for (let i = 0; i < nodes.length; i++) {
		const node = nodes[i] as FormulaAst;
		const copy = copyTree(node, t, rows);
		if (copy !== node && !out) out = nodes.slice(0, i);
		out?.push(copy);
	}
	return out ?? nodes;
}

/** The tree of a formula shaped like the template, or undefined when it must be parsed itself. */
export function fromTemplate(t: Template, formula: string, shape: Shape): FormulaAst | undefined {
	const rows = rowsFor(t, formula, shape);
	return rows ? copyTree(t.ast, t, rows) : undefined;
}
