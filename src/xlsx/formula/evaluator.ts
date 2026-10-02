// Evaluates a formula syntax tree to a value.
import type { DefinedName } from '../model.js';
import { FormulaError, type FormulaAst, normalizeLocalName } from './ast.js';
import { callFunction } from './call.js';
import type { EvalHost, Frame } from './context.js';
import { broadcast1, broadcast2, scalarBinary, scalarUnary } from './operators.js';
import {
	implicitIntersection,
	intersectOperator,
	rangeOperator,
	resolveRefNode,
	resolveStructured,
	sheetIndex,
	toMatrix,
	unionOperator,
} from './references.js';
import {
	err,
	ERR,
	ErrorSignal,
	isError,
	LambdaValue,
	Matrix,
	RefValue,
	type Scalar,
	type Scope,
	type Value,
} from './values.js';

const MAX_NAME_DEPTH = 32;

/** Evaluates `node`; references stay references (the caller decides how to read them). */
export function evaluateNode(node: FormulaAst, frame: Frame): Value {
	switch (node.type) {
		case 'number':
		case 'string':
		case 'boolean':
			return node.value;
		case 'error':
			return err(node.code);
		case 'missing':
			return null;
		case 'array':
			return new Matrix(
				node.rows.map((row) =>
					row.map((item) => (typeof item === 'object' ? err(item.error) : item)),
				),
			);
		case 'ref':
			return resolveRefNode(node, frame);
		case 'name':
			return resolveName(node, frame);
		case 'structured':
			return resolveStructured(node.ref, frame);
		case 'unary': {
			if (node.op === '+') return evaluateNode(node.operand, frame);
			const value = evaluateNode(node.operand, frame);
			if (node.op === '@') return implicitIntersection(value, frame);
			return broadcast1(operand(value, frame), (x) => scalarUnary('-', x));
		}
		case 'percent':
			return broadcast1(operand(evaluateNode(node.operand, frame), frame), (x) =>
				scalarUnary('%', x),
			);
		case 'binary':
			return evaluateBinary(node, frame);
		case 'call':
			return callFunction(node, frame);
		case 'invoke': {
			const callee = evaluateNode(node.callee, frame);
			const args = node.args.map((arg) => evaluateNode(arg, frame));
			return invokeLambda(callee, args, frame);
		}
	}
}

/** Dereferences an operand: one cell becomes its value, larger ranges a matrix. */
export function operand(value: Value, frame: Frame): Scalar | Matrix {
	if (value instanceof RefValue) {
		if (frame.legacy && !value.isCell()) return implicitIntersection(value, frame);
		if (value.isCell()) {
			const area = value.areas[0];
			return area
				? frame.host.readCell(area.sheet, area.range.start.row, area.range.start.col)
				: null;
		}
		try {
			return toMatrix(frame.host, value);
		} catch (e) {
			if (e instanceof ErrorSignal) return e.value;
			throw e;
		}
	}
	if (value instanceof LambdaValue) return ERR.CALC;
	return value;
}

function evaluateBinary(node: Extract<FormulaAst, { type: 'binary' }>, frame: Frame): Value {
	const left = evaluateNode(node.left, frame);
	if (node.op === ':' || node.op === ' ' || node.op === ',') {
		if (isError(left)) return left;
		const right = evaluateNode(node.right, frame);
		if (isError(right)) return right;
		if (node.op === ':') return rangeOperator(left, right);
		if (node.op === ' ') return intersectOperator(left, right);
		return unionOperator(left, right);
	}
	const right = evaluateNode(node.right, frame);
	const snap = node === frame.root && (node.op === '+' || node.op === '-');
	return broadcast2(operand(left, frame), operand(right, frame), (a, b) =>
		scalarBinary(node.op, a, b, snap),
	);
}

/** The defined name visible from `sheet` (a sheet-local name wins over a global one). */
export function findDefinedName(
	host: EvalHost,
	name: string,
	sheet: number,
	explicitSheet: boolean,
): DefinedName | undefined {
	const upper = name.toUpperCase();
	let global: DefinedName | undefined;
	for (const defined of host.workbook.definedNames) {
		if (defined.name.toUpperCase() !== upper) continue;
		if (defined.localSheet === sheet) return defined;
		if (defined.localSheet === undefined) global = defined;
	}
	return explicitSheet && !global ? undefined : global;
}

/** Evaluates a defined name's formula, or `undefined` when there is no such name. */
export function evaluateDefinedName(
	name: string,
	sheet: number,
	explicitSheet: boolean,
	frame: Frame,
): Value | undefined {
	const defined = findDefinedName(frame.host, name, sheet, explicitSheet);
	if (!defined) return undefined;
	if (frame.depth >= MAX_NAME_DEPTH) return ERR.NAME;
	const ast = frame.host.parse(defined.formula);
	if (ast instanceof FormulaError) return ERR.NAME;
	return evaluateNode(ast, { ...frame, depth: frame.depth + 1, scope: undefined, root: undefined });
}

function resolveName(node: Extract<FormulaAst, { type: 'name' }>, frame: Frame): Value {
	if (!node.prefix) {
		const local = frame.scope?.get(normalizeLocalName(node.name));
		if (local !== undefined) return local;
	}
	let sheet = frame.sheet;
	if (node.prefix) {
		sheet = sheetIndex(frame.host, node.prefix.sheet);
		if (sheet < 0 || node.prefix.book !== undefined) return ERR.REF;
	}
	const value = evaluateDefinedName(node.name, sheet, node.prefix !== undefined, frame);
	if (value !== undefined) return value;
	if (!node.prefix) {
		const table = resolveStructured({ table: node.name, specials: [] }, frame);
		if (!(isError(table) && table.error === '#NAME?')) return table;
	}
	return ERR.NAME;
}

/** Calls a LAMBDA closure with already evaluated arguments. */
export function invokeLambda(callee: Value, args: Value[], frame: Frame): Value {
	if (isError(callee)) return callee;
	if (!(callee instanceof LambdaValue)) return ERR.VALUE;
	if (args.length > callee.params.length) return ERR.VALUE;
	const scope = new Map<string, Value>(callee.scope ?? []);
	callee.params.forEach((param, i) => scope.set(param, args[i] ?? null));
	return evaluateNode(callee.body, { ...frame, scope, root: undefined });
}

/** Extends a scope with one variable. */
export function extendScope(scope: Scope | undefined, name: string, value: Value): Scope {
	const next = new Map<string, Value>(scope ?? []);
	next.set(normalizeLocalName(name), value);
	return next;
}
