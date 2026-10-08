// Parsed formulas by text and by shape (see formula-shape.ts): building the dependency graph of a
// workbook with tens of thousands of formulas tokenizes only one copy of each filled formula.
import type { FormulaAst, FormulaError } from './ast';
import {
	formulaShape,
	fromTemplate,
	makeTemplate,
	type Shape,
	type Template,
} from './formula-shape';

/** Formulas longer than this are parsed on their own (their trees may be deep). */
const MAX_TEMPLATE_LENGTH = 2048;
const MAX_ENTRIES = 100_000;
const MAX_TEMPLATES = 10_000;
const NO_RUNS: Shape = { hash: 0, starts: [], ends: [] };

const treeOf = (made: Template | FormulaError): FormulaAst | FormulaError =>
	'ast' in made ? made.ast : made;

/** Parsed formulas, keyed by text and, for reuse across copies, by shape. */
export class ParseCache {
	private readonly byText = new Map<string, FormulaAst | FormulaError>();
	private readonly templates = new Map<number, Template>();

	parse(formula: string): FormulaAst | FormulaError {
		let ast = this.byText.get(formula);
		if (ast) return ast;
		ast = this.parseNew(formula);
		if (this.byText.size > MAX_ENTRIES) this.byText.clear();
		this.byText.set(formula, ast);
		return ast;
	}

	private parseNew(formula: string): FormulaAst | FormulaError {
		if (formula.length > MAX_TEMPLATE_LENGTH) return treeOf(makeTemplate(formula, NO_RUNS));
		const shape = formulaShape(formula);
		const template = this.templates.get(shape.hash);
		const reused = template && fromTemplate(template, formula, shape);
		if (reused) return reused;
		const made = makeTemplate(formula, shape);
		if (!template && 'ast' in made) {
			if (this.templates.size > MAX_TEMPLATES) this.templates.clear();
			this.templates.set(shape.hash, made);
		}
		return treeOf(made);
	}
}
