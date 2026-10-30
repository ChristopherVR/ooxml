import {
	evaluateVisioFormula,
	type VisioFormulaAst,
	type VisioFormulaReference,
	type VisioFormulaValue,
} from './formula';

/**
 * The text a Menu formula yields, for the shapes Visio ships: a string, or `IF(condition, text,
 * text)` over cells, nested and joined with `&`. Anything else is not text this editor can
 * compute and gives `undefined`.
 */
export function evaluateMenuText(
	ast: VisioFormulaAst,
	read: (reference: VisioFormulaReference) => VisioFormulaValue,
): string | undefined {
	if (ast.kind === 'string') return ast.value;
	if (ast.kind === 'number') return String(ast.value);
	if (ast.kind === 'binary' && ast.operator === '&') {
		const left = evaluateMenuText(ast.left, read),
			right = evaluateMenuText(ast.right, read);
		return left === undefined || right === undefined ? undefined : left + right;
	}
	if (ast.kind === 'call' && ast.name === 'IF' && ast.args.length === 3) {
		const condition = evaluateVisioFormula(ast.args[0]!, read, { bareLengths: true });
		return evaluateMenuText(condition.value !== 0 ? ast.args[1]! : ast.args[2]!, read);
	}
	return undefined;
}
