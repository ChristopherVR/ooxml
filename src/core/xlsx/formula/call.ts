// Function calls: argument evaluation, dynamic-array lifting and the context implementations see.
import { FormulaError, type FormulaAst, normalizeLocalName } from './ast.js';
import type { CallContext, Frame, LazyArg } from './context.js';
import { currentDate1904 } from './date-serial.js';
import { evaluateDefinedName, evaluateNode, invokeLambda } from './evaluator.js';
import { getFunction } from './functions/registry.js';
import { type FunctionSpec, paramKind } from './functions/types.js';
import { broadcastShape, pick } from './operators.js';
import { parseFormula } from './parser.js';
import { implicitIntersection, toMatrix } from './references.js';
import {
	ERR,
	ErrorSignal,
	LambdaValue,
	Matrix,
	RefValue,
	type Scalar,
	type Scope,
	type Value,
} from './values.js';

const guard = (fn: () => Value): Value => {
	try {
		return fn();
	} catch (e) {
		if (e instanceof ErrorSignal) return e.value;
		throw e;
	}
};

/** Evaluates a function call node. */
export function callFunction(node: Extract<FormulaAst, { type: 'call' }>, frame: Frame): Value {
	const local = frame.scope?.get(normalizeLocalName(node.rawName));
	if (local instanceof LambdaValue) {
		return invokeLambda(
			local,
			node.args.map((a) => evaluateNode(a, frame)),
			frame,
		);
	}
	const spec = getFunction(node.name);
	if (!spec) {
		const named = evaluateDefinedName(node.rawName, frame.sheet, false, frame);
		if (named instanceof LambdaValue) {
			return invokeLambda(
				named,
				node.args.map((a) => evaluateNode(a, frame)),
				frame,
			);
		}
		return ERR.NAME;
	}
	if (node.args.length < spec.minArgs || node.args.length > spec.maxArgs) return ERR.VALUE;
	const ctx = createContext(frame);
	if (spec.lazy) {
		const lazy = spec.lazy;
		const args: LazyArg[] = node.args.map((arg) => ({
			node: arg,
			get: () => evaluateNode(arg, frame),
		}));
		return guard(() => lazy(args, ctx));
	}
	const arrayFrame = frame.legacy ? { ...frame, legacy: false } : frame;
	const args = node.args.map((arg, i) =>
		evaluateNode(arg, paramKind(spec, i) === 'any' ? arrayFrame : frame),
	);
	return applyFunction(spec, args, ctx);
}

/** Calls a function implementation, lifting it over arrays passed to `value` parameters. */
export function applyFunction(spec: FunctionSpec, args: Value[], ctx: CallContext): Value {
	const fn = spec.fn;
	if (!fn) return ERR.VALUE;
	const lifted: Matrix[] = [];
	for (let i = 0; i < args.length; i++) {
		if (paramKind(spec, i) !== 'value') continue;
		let arg = args[i] as Value;
		if (arg instanceof RefValue) {
			if (ctx.frame.legacy && !arg.isCell()) arg = ctx.toScalar(arg);
			else if (arg.isCell()) {
				const area = arg.areas[0];
				arg = area ? ctx.readCell(area.sheet, area.range.start.row, area.range.start.col) : null;
			} else {
				try {
					arg = ctx.toMatrix(arg);
				} catch (e) {
					if (e instanceof ErrorSignal) return e.value;
					throw e;
				}
			}
		}
		if (arg instanceof Matrix) {
			if (arg.rows === 1 && arg.cols === 1) arg = arg.get(0, 0);
			else lifted.push(arg);
		}
		args[i] = arg;
	}
	if (lifted.length === 0) return guard(() => fn(args, ctx));
	return broadcastShape(lifted, (r, c) => {
		const element = args.map((arg, i) =>
			arg instanceof Matrix && paramKind(spec, i) === 'value' ? pick(arg, r, c) : arg,
		);
		return toElement(
			guard(() => fn(element, ctx)),
			ctx,
		);
	});
}

function toElement(value: Value, ctx: CallContext): Scalar {
	if (value instanceof Matrix) return value.get(0, 0);
	if (value instanceof RefValue) return ctx.toScalar(value);
	if (value instanceof LambdaValue) return ERR.CALC;
	return value;
}

/** Builds the context passed to function implementations. */
export function createContext(frame: Frame): CallContext {
	const host = frame.host;
	return {
		frame,
		workbook: host.workbook,
		sheet: frame.sheet,
		row: frame.row,
		col: frame.col,
		date1904: currentDate1904(),
		toMatrix: (value) => toMatrix(host, value),
		toScalar: (value) => implicitIntersection(value, frame),
		forEach(value, visit) {
			if (value instanceof RefValue) {
				for (const area of value.areas) {
					host.forEachStored(area.sheet, area.range, (v) => visit(v, 'ref'));
				}
				return;
			}
			if (value instanceof Matrix) {
				value.forEachValue((v) => visit(v, 'array'));
				return;
			}
			if (value instanceof LambdaValue) throw new ErrorSignal(ERR.CALC);
			visit(value, 'direct');
		},
		readCell: (sheet, row, col) => host.readCell(sheet, row, col),
		evaluate: (node, scope?: Scope) =>
			evaluateNode(
				node,
				scope ? { ...frame, scope, root: undefined } : { ...frame, root: undefined },
			),
		callLambda: (fn, args) => invokeLambda(fn, args, frame),
		parseReference(text) {
			let ast: FormulaAst;
			try {
				ast = parseFormula(text.trim());
			} catch (e) {
				if (e instanceof FormulaError) return undefined;
				throw e;
			}
			if (ast.type !== 'ref' && ast.type !== 'name' && ast.type !== 'binary') return undefined;
			if (ast.type === 'binary' && ast.op !== ':') return undefined;
			const value = evaluateNode(ast, { ...frame, root: undefined });
			return value instanceof RefValue ? value : undefined;
		},
	};
}
