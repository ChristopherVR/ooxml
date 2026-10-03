import { themeInputs } from './formula-theme.js';
import { formulaFailure, formulaLimit, parseVisioFormula } from './formula.js';
import type {
	VisioFormulaAst,
	VisioFormulaLimits,
	VisioFormulaReference,
	VisioFormulaValue,
} from './formula.js';

const supported = new Set([
	'GUARD',
	'IF',
	'ABS',
	'MIN',
	'MAX',
	'SQRT',
	'SIN',
	'COS',
	'TAN',
	'ATAN2',
	'SIGN',
	'PI',
]);
// Known pure numeric functions may be independent of an edit even when evaluation is unsupported.
const staticUnsupported = new Set([
	'ROUND',
	'INT',
	'MOD',
	'POW',
	'EXP',
	'LN',
	'LOG',
	'LOG10',
	'ASIN',
	'ACOS',
	'ATAN',
	'CEILING',
	'FLOOR',
	'TRUNC',
	'AND',
	'OR',
	'NOT',
	'SUM',
	'RGB',
	'HSL',
	'THEMEGUARD',
	'THEMEVAL',
	'THEME',
	'ISTHEMED',
	'BITAND',
	'BITOR',
	'BITXOR',
	'TINT',
	'MSOTINT',
	'SHADE',
	'LUMDIFF',
	'MODULUS',
]);
// These functions can introduce references that cannot be identified from argument ASTs.
const dynamic = new Set([
	'INDIRECT',
	'EVALCELL',
	'EVALTEXT',
	'DEPENDSON',
	'SETATREF',
	'SETATREFEXPR',
	'SETATREFEVAL',
	'GETREF',
	'REF',
	'SHAPETEXT',
	'CALLOUTTARGET',
	'CONTAINERSHEETREF',
	'MEMBERSHEETREF',
	'PARENT',
	'PAR',
]);
export interface VisioFormulaAnalysis {
	references: VisioFormulaReference[];
	unsupportedFunctions: string[];
	dynamic: boolean;
	guarded: boolean;
}
function traverse(
	ast: VisioFormulaAst,
	limits: VisioFormulaLimits,
	visit: (node: VisioFormulaAst) => void,
): void {
	const maxNodes = formulaLimit(limits.maxNodes, 1024),
		maxDepth = formulaLimit(limits.maxDepth, 64);
	const queue: [VisioFormulaAst, number][] = [[ast, 0]];
	let count = 0;
	while (queue.length) {
		const [node, depth] = queue.pop()!;
		if (++count > maxNodes || depth > maxDepth)
			return formulaFailure('limit', 'Formula analysis limit');
		visit(node);
		if (node.kind === 'unary') queue.push([node.operand, depth + 1]);
		if (node.kind === 'binary') queue.push([node.left, depth + 1], [node.right, depth + 1]);
		if (node.kind === 'call') for (const arg of node.args) queue.push([arg, depth + 1]);
	}
}
export function analyzeVisioFormula(
	source: string | VisioFormulaAst,
	limits: VisioFormulaLimits = {},
): VisioFormulaAnalysis {
	const ast = typeof source === 'string' ? parseVisioFormula(source, limits) : source;
	const references = new Map<string, VisioFormulaReference>(),
		unsupportedFunctions = new Set<string>();
	let unsafe = false,
		guarded = false;
	traverse(ast, limits, (node) => {
		if (node.kind === 'reference')
			references.set(
				`${node.reference.shapeId ?? ''}!${node.reference.cell.toLowerCase()}`,
				node.reference,
			);
		if (node.kind === 'call') {
			const inputs =
				node.name === 'ISTHEMED'
					? ['ColorSchemeIndex']
					: ['THEMEVAL', 'THEME'].includes(node.name)
						? themeInputs
						: [];
			for (const cell of inputs) references.set(`!${cell.toLowerCase()}`, { cell });
			if (!supported.has(node.name)) unsupportedFunctions.add(node.name);
			if (
				dynamic.has(node.name) ||
				(!supported.has(node.name) && !staticUnsupported.has(node.name))
			)
				unsafe = true;
			if (node.name === 'GUARD') guarded = true;
		}
	});
	return {
		references: [...references.values()],
		unsupportedFunctions: [...unsupportedFunctions],
		dynamic: unsafe,
		guarded,
	};
}
function finite(value: VisioFormulaValue): VisioFormulaValue {
	if (!Number.isFinite(value.value)) return formulaFailure('value', 'Formula result is not finite');
	if (!['scalar', 'length', 'angle', 'time'].includes(value.unit))
		return formulaFailure('unit', 'Unknown resolved unit');
	return value;
}
function compatible(a: VisioFormulaValue, b: VisioFormulaValue): VisioFormulaValue['unit'] {
	if (a.unit === b.unit) return a.unit;
	// The dimensionless zero is valid in comparisons and sums with dimensional values.
	if (a.unit === 'scalar' && a.value === 0) return b.unit;
	if (b.unit === 'scalar' && b.value === 0) return a.unit;
	return formulaFailure('unit', `Incompatible units ${a.unit} and ${b.unit}`);
}
function binary(operator: string, a: VisioFormulaValue, b: VisioFormulaValue): VisioFormulaValue {
	let unit: VisioFormulaValue['unit'];
	if (operator === '+' || operator === '-') {
		unit = compatible(a, b);
		return finite({ value: operator === '+' ? a.value + b.value : a.value - b.value, unit });
	}
	if (operator === '*') {
		if (a.unit !== 'scalar' && b.unit !== 'scalar')
			return formulaFailure('unit', 'Compound dimensions are unsupported');
		return finite({ value: a.value * b.value, unit: a.unit === 'scalar' ? b.unit : a.unit });
	}
	if (operator === '/') {
		if (b.value === 0) return formulaFailure('value', 'Division by zero');
		if (b.unit !== 'scalar' && a.unit !== b.unit)
			return formulaFailure('unit', 'Inverse dimensions are unsupported');
		return finite({ value: a.value / b.value, unit: a.unit === b.unit ? 'scalar' : a.unit });
	}
	if (operator === '^') {
		if (b.unit !== 'scalar' || (a.unit !== 'scalar' && b.value !== 1 && b.value !== 0))
			return formulaFailure('unit', 'Dimensional powers are unsupported');
		return finite({ value: a.value ** b.value, unit: b.value === 0 ? 'scalar' : a.unit });
	}
	compatible(a, b);
	const comparisons: Record<string, boolean> = {
		'=': a.value === b.value,
		'<>': a.value !== b.value,
		'!=': a.value !== b.value,
		'<': a.value < b.value,
		'>': a.value > b.value,
		'<=': a.value <= b.value,
		'>=': a.value >= b.value,
	};
	if (!(operator in comparisons))
		return formulaFailure('syntax', `Unsupported operator ${operator}`);
	return { value: comparisons[operator] ? 1 : 0, unit: 'scalar' };
}
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
		if (node.kind === 'binary')
			return binary(node.operator, evaluate(node.left, depth + 1), evaluate(node.right, depth + 1));
		const name = node.name;
		if (!supported.has(name)) return formulaFailure('unsupported', `Unsupported function ${name}`);
		const arity = (min: number, max = min) => {
			if (node.args.length < min || node.args.length > max)
				return formulaFailure('arity', `Invalid ${name} argument count`);
		};
		const arg = (index: number) => evaluate(node.args[index]!, depth + 1);
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
			if (value.unit !== 'scalar') return formulaFailure('unit', 'Dimensional SQRT unsupported');
			return finite({ value: Math.sqrt(value.value), unit: 'scalar' });
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
