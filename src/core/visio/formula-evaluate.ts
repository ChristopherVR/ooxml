import { supported, traverse } from './formula-analyze';
import { evaluateBound } from './formula-bound';
import { binaryOperator as binary, evaluatedShapeFrame } from './formula-operators';
import {
	finiteFormulaValue as finite,
	compatibleFormulaUnits as compatible,
	formulaProductUnit,
	formulaQuotientUnit,
	formulaExponentUnit,
	formulaSquareRootUnit,
} from './formula-arithmetic';
import {
	numericFormulaFunctions,
	evaluateNumericFormulaFunction,
} from './formula-numeric-functions';
import {
	evaluateFormulaText,
	evaluateTextNumberFunction,
	textNumberFunctions,
	textResultFunctions,
} from './formula-text-functions';
import {
	angleToLocal,
	localToParent,
	visioDateSerial,
	type ShapeFrame,
} from './formula-extra-functions';
import { mathFormulaFunctions, evaluateMathFormulaFunction } from './formula-math-functions';
import { formulaFailure, formulaLimit, parseVisioFormula, VisioFormulaError } from './formula';
import type {
	VisioFormulaAst,
	VisioFormulaLimits,
	VisioFormulaReference,
	VisioFormulaValue,
} from './formula';

export function evaluateVisioFormula(
	source: string | VisioFormulaAst,
	resolve: (reference: VisioFormulaReference) => VisioFormulaValue,
	limits: VisioFormulaLimits = {},
): VisioFormulaValue {
	const ast = typeof source === 'string' ? parseVisioFormula(source, limits) : source;
	const maxSteps = formulaLimit(limits.maxSteps, 8192),
		maxDepth = formulaLimit(limits.maxDepth, 64);
	// Validate caller-supplied ASTs as well as parsed ones before recursive evaluation.
	traverse(ast, limits, () => {});
	let steps = 0;
	const shapeFrame = (sheet: VisioFormulaAst) => evaluatedShapeFrame(sheet, resolve);
	const evaluate = (node: VisioFormulaAst, depth = 0): VisioFormulaValue => {
		if (++steps > maxSteps || depth > maxDepth)
			return formulaFailure('limit', 'Formula evaluation limit');
		limits.onStep?.();
		if (node.kind === 'number') return finite({ value: node.value, unit: node.unit });
		if (node.kind === 'string')
			return formulaFailure('value', 'String results are outside the numeric formula subset');
		if (node.kind === 'reference') return finite(resolve(node.reference));
		if (node.kind === 'unary') {
			const result = evaluate(node.operand, depth + 1);
			return finite({
				value: node.operator === '-' ? -result.value : result.value,
				unit: result.unit,
			});
		}
		if (node.kind === 'binary') {
			let left = evaluate(node.left, depth + 1),
				right = evaluate(node.right, depth + 1);
			if (limits.bareLengths && (node.operator === '+' || node.operator === '-')) {
				if (left.unit === 'length' && right.unit === 'scalar' && node.right.kind === 'number')
					right = { ...right, unit: 'length' };
				else if (right.unit === 'length' && left.unit === 'scalar' && node.left.kind === 'number')
					left = { ...left, unit: 'length' };
			}
			return binary(node.operator, left, right);
		}
		const name = node.name;
		if ((name === 'TEXTWIDTH' || name === 'TEXTHEIGHT') && limits.text) {
			const text = node.args[0];
			const count = node.args.length;
			if (
				text?.kind !== 'reference' ||
				text.reference.shapeId !== undefined ||
				text.reference.cell.toLowerCase() !== 'thetext' ||
				count > 2 ||
				(name === 'TEXTHEIGHT' && count !== 2)
			)
				return formulaFailure('unsupported', `Unsupported ${name} arguments`);
			const limit = count === 2 ? evaluate(node.args[1]!, depth + 1) : undefined;
			if (limit && limit.unit !== 'length' && limit.unit !== 'scalar')
				return formulaFailure('unit', `${name} requires a length`);
			const value =
				name === 'TEXTWIDTH' ? limits.text.width(limit?.value) : limits.text.height(limit!.value);
			if (value === undefined)
				return formulaFailure('unsupported', 'The text cannot be measured reliably');
			return finite({ value, unit: 'length' });
		}
		if (!supported.has(name)) return formulaFailure('unsupported', `Unsupported function ${name}`);
		const arity = (min: number, max = min) => {
			if (node.args.length < min || node.args.length > max)
				return formulaFailure('arity', `Invalid ${name} argument count`);
		};
		const arg = (index: number) => evaluate(node.args[index]!, depth + 1);
		if (name === 'PAGENUMBER' || name === 'PAGECOUNT') {
			arity(0);
			const value = name === 'PAGENUMBER' ? limits.pageNumber : limits.pageCount;
			if (value === undefined || !Number.isSafeInteger(value) || value < 0)
				return formulaFailure('context', `${name} requires a valid document context`);
			return { value, unit: 'scalar' };
		}
		const textContext = {
			number: (child: VisioFormulaAst) => evaluate(child, depth + 1),
			...(limits.textOf ? { textOf: limits.textOf } : {}),
			tick: () => {
				if (++steps > maxSteps) formulaFailure('limit', 'Formula evaluation limit');
			},
		};
		if (textNumberFunctions.has(name))
			return evaluateTextNumberFunction(name, node.args, textContext);
		if (name === 'INDIRECT') {
			// INDIRECT(address): the value of the cell a text address names, read from this shape.
			arity(1);
			const target = parseVisioFormula(evaluateFormulaText(node.args[0]!, textContext, 1));
			if (target.kind !== 'reference')
				return formulaFailure('value', 'INDIRECT needs a cell address');
			return finite(resolve(target.reference));
		}
		if (name === 'DATE') {
			// DATE(year, month, day): an OLE date serial; the unit keeps it from posing as a plain number.
			arity(3);
			const [year, month, day] = node.args.map((_, index) => {
				const part = arg(index);
				if (part.unit !== 'scalar') return formulaFailure('unit', 'DATE requires plain numbers');
				return part.value;
			});
			return { value: visioDateSerial(year!, month!, day!), unit: 'time' };
		}
		if (name === 'SHEETREF' || name === 'LOCTOPAR')
			return formulaFailure('unsupported', `${name} does not return a number`);
		if (name === 'ANGLETOLOC') {
			arity(2);
			const angle = arg(1);
			if (angle.unit !== 'angle' && angle.unit !== 'scalar')
				return formulaFailure('unit', 'ANGLETOLOC requires an angle');
			return { value: angleToLocal(shapeFrame(node.args[0]!), angle.value), unit: 'angle' };
		}
		if (textResultFunctions.has(name) || name === 'PNT')
			return formulaFailure('unsupported', `${name} does not return a number`);
		if (name === 'PNTX' || name === 'PNTY') {
			// PNT(x, y) builds a point; its coordinates are read straight back with PNTX / PNTY.
			arity(1);
			const point = node.args[0]!;
			if (point.kind === 'call' && point.name === 'LOCTOPAR') {
				// LOCTOPAR(SHEETREF(), x, y): the point in the parent's coordinates.
				if (point.args.length !== 3)
					return formulaFailure('arity', 'Invalid LOCTOPAR argument count');
				const [x, y] = [1, 2].map((index) => {
					const part = evaluate(point.args[index]!, depth + 2);
					if (part.unit !== 'scalar' && part.unit !== 'length')
						return formulaFailure('unit', 'LOCTOPAR requires lengths');
					return part.value;
				});
				const parent = localToParent(shapeFrame(point.args[0]!), x!, y!);
				return finite({ value: parent[name === 'PNTX' ? 0 : 1], unit: 'length' });
			}
			if (point.kind !== 'call' || point.name !== 'PNT')
				return formulaFailure('unsupported', `${name} of something other than PNT(x,y)`);
			if (point.args.length !== 2) return formulaFailure('arity', 'Invalid PNT argument count');
			const [x, y] = point.args.map((coordinate) => evaluate(coordinate, depth + 2));
			const picked = name === 'PNTX' ? x! : y!;
			for (const item of [x!, y!])
				if (item.unit !== 'scalar' && item.unit !== 'length')
					return formulaFailure('unit', 'PNT requires lengths');
			return finite({ value: picked.value, unit: 'length' });
		}
		if (name === 'ISERROR' || name === 'IFERROR') {
			arity(name === 'ISERROR' ? 1 : 2);
			try {
				const result = arg(0);
				return name === 'ISERROR' ? { value: 0, unit: 'scalar' } : result;
			} catch (error) {
				// Only a value error (#VALUE!, #DIV/0!) is an error to the sheet; a function this
				// evaluator cannot run, or a limit, must still refuse the formula.
				if (!(error instanceof VisioFormulaError) || error.code !== 'value') throw error;
				return name === 'ISERROR' ? { value: 1, unit: 'scalar' } : arg(1);
			}
		}
		if (numericFormulaFunctions.has(name))
			return evaluateNumericFormulaFunction(
				name,
				node.args.map((_, index) => arg(index)),
			);
		if (mathFormulaFunctions.has(name))
			return evaluateMathFormulaFunction(
				name,
				node.args.map((_, index) => arg(index)),
			);
		if (name === 'PI') {
			arity(0);
			return { value: Math.PI, unit: 'scalar' };
		}
		if (name === 'GUARD') {
			arity(1);
			return arg(0);
		}
		if (name === 'IF') {
			arity(3);
			const condition = arg(0);
			if (condition.unit !== 'scalar') return formulaFailure('unit', 'IF condition must be scalar');
			return arg(condition.value === 0 ? 2 : 1);
		}
		if (name === 'MIN' || name === 'MAX') {
			arity(1, 1024);
			const args = node.args.map((_, i) => arg(i));
			let result = args[0]!;
			for (const item of args.slice(1)) {
				const unit = compatible(result, item);
				result = {
					value:
						name === 'MIN'
							? Math.min(result.value, item.value)
							: Math.max(result.value, item.value),
					unit,
				};
			}
			return finite(result);
		}
		if (name === 'BOUND') return evaluateBound(node.args, arg);
		if (name === 'ATAN2') {
			arity(2);
			const y = arg(0),
				x = arg(1);
			compatible(y, x);
			return { value: Math.atan2(y.value, x.value), unit: 'angle' };
		}
		if (name === 'SIGN') {
			arity(1, 2);
			const value = arg(0);
			const fuzz = node.args.length === 2 ? arg(1) : { value: 1e-9, unit: 'scalar' as const };
			if (fuzz.value < 0 || (fuzz.unit !== 'scalar' && fuzz.unit !== value.unit))
				return formulaFailure('unit', 'Invalid SIGN fuzz');
			return {
				value: Math.abs(value.value) <= fuzz.value ? 0 : Math.sign(value.value),
				unit: 'scalar',
			};
		}
		arity(1);
		const value = arg(0);
		if (name === 'ABS') return { value: Math.abs(value.value), unit: value.unit };
		if (name === 'SQRT') {
			return finite({ value: Math.sqrt(value.value), unit: formulaSquareRootUnit(value.unit) });
		}
		if (value.unit !== 'angle' && value.unit !== 'scalar')
			return formulaFailure('unit', `${name} requires radians`);
		return finite({
			value:
				name === 'SIN'
					? Math.sin(value.value)
					: name === 'COS'
						? Math.cos(value.value)
						: Math.tan(value.value),
			unit: 'scalar',
		});
	};
	return evaluate(ast);
}

export { analyzeVisioFormula, isClassifiedVisioFunction } from './formula-analyze';
export type { VisioFormulaAnalysis } from './formula-analyze';
