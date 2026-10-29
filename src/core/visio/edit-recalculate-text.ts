import { analyzeVisioFormula, parseVisioFormula, type VisioFormulaAst } from './formula';
import { isTheText, textNumberFunctions, textResultFunctions } from './formula-text-functions';
import { decodeVisioPlainText } from './plain-text';
import { attribute, children } from './sheet';

/**
 * The plain text a local shape shows, for formulas that read `TheText`. Unknown (undefined) when
 * the text is inherited from a master, holds fields or rich runs, or the shape is not on the page.
 */
export function plainShapeText(root: Element | undefined, shapeId: string): string | undefined {
	if (!root) return undefined;
	const pending = children(children(root, 'Shapes')[0], 'Shape');
	while (pending.length) {
		const shape = pending.pop()!;
		if (attribute(shape, 'ID') === shapeId) {
			if (shape.hasAttribute('Master') || shape.hasAttribute('MasterShape')) return undefined;
			if (children(shape, 'Section').some((section) => attribute(section, 'N') === 'Field'))
				return undefined;
			const texts = children(shape, 'Text');
			if (texts.length > 1) return undefined;
			const node = texts[0];
			if (!node) return '';
			const plain = Array.from(node.childNodes).every(
				(item) => item.nodeType === 3 || item.nodeType === 4,
			);
			return plain ? decodeVisioPlainText(node.textContent ?? '') : undefined;
		}
		pending.push(...children(children(shape, 'Shapes')[0], 'Shape'));
	}
	return undefined;
}

const TEXT_READERS = new Set([...textNumberFunctions, ...textResultFunctions]);

/**
 * A formula whose every read of a shape's text is an argument of SHAPETEXT or a text function
 * (`LEN(SHAPETEXT(TheText))`, `FIND("x",TheText)`), so it can be computed again from the new text.
 * A bare `TheText` in arithmetic is not one.
 */
export function textReadFormula(source: string): boolean {
	try {
		const ast = parseVisioFormula(source);
		const analysis = analyzeVisioFormula(ast, { shapeTextReads: true });
		if (analysis.dynamic || analysis.unsupportedFunctions.length) return false;
		// A text result cannot be written into a numeric cache.
		const textResult =
			ast.kind === 'string' ||
			isTheText(ast) ||
			(ast.kind === 'binary' && ast.operator === '&') ||
			(ast.kind === 'call' && textResultFunctions.has(ast.name));
		if (textResult) return false;
		let reads = 0;
		const walk = (node: VisioFormulaAst, parent?: VisioFormulaAst): boolean => {
			if (isTheText(node)) {
				reads++;
				return parent?.kind === 'call' && TEXT_READERS.has(parent.name);
			}
			if (node.kind === 'unary') return walk(node.operand, node);
			if (node.kind === 'binary') return walk(node.left, node) && walk(node.right, node);
			if (node.kind === 'call') return node.args.every((arg) => walk(arg, node));
			return true;
		};
		return walk(ast) && reads > 0;
	} catch {
		return false;
	}
}
