// Evaluate Formula: reduces a formula one sub-expression at a time. The next sub-expression is the
// leftmost innermost one whose operands are already values; Evaluate computes it with the shared
// calc engine and puts the result in its place. Step In opens the formula of a referenced cell,
// Step Out returns to the formula that referenced it.
import { formatAddress } from '../address';
import { quoteSheetName } from '../address';
import { FormulaError, type FormulaAst } from '../formula/ast';
import { parseFormula } from '../formula/parser';
import { isCellError, type CellValue, type Workbook } from '../model';
import { printFormula } from './ast-text';

export interface StepEvaluator {
	/** Evaluates a formula (no leading `=`) as if it were stored at `at`, keeping arrays. */
	evaluateArray(formula: string, at: { sheet: number; row: number; col: number }): CellValue[][];
}

interface Frame {
	sheet: number;
	row: number;
	col: number;
	ast: FormulaAst;
}

export interface StepView {
	/** `Sheet1!$A$1` of the formula being evaluated. */
	reference: string;
	/** The formula as reduced so far, with its leading `=`. */
	text: string;
	/** Offsets into `text` of the sub-expression Evaluate will compute next (none when finished). */
	underline?: { start: number; end: number };
	finished: boolean;
	canStepIn: boolean;
	canStepOut: boolean;
}

const LAZY = new Set(['IF', 'IFERROR', 'IFNA']);
const isLiteral = (n: FormulaAst): boolean =>
	n.type === 'number' ||
	n.type === 'string' ||
	n.type === 'boolean' ||
	n.type === 'error' ||
	n.type === 'array' ||
	n.type === 'missing';

function childrenOf(node: FormulaAst): FormulaAst[] {
	switch (node.type) {
		case 'unary':
		case 'percent':
			return [node.operand];
		case 'binary':
			return [node.left, node.right];
		case 'call':
			return node.args;
		case 'invoke':
			return [node.callee, ...node.args];
		default:
			return [];
	}
}

/** The next node to compute: first in evaluation order whose operands are all values. */
function nextNode(node: FormulaAst): FormulaAst | undefined {
	if (isLiteral(node)) return undefined;
	// IF/IFERROR/IFNA only compute the branch they need once the first argument is a value.
	if (node.type === 'call' && LAZY.has(node.name) && node.args[0] && isLiteral(node.args[0]))
		return node;
	for (const child of childrenOf(node)) {
		const found = nextNode(child);
		if (found) return found;
	}
	return node;
}

function replaceNode(root: FormulaAst, target: FormulaAst, by: FormulaAst): FormulaAst {
	if (root === target) return by;
	switch (root.type) {
		case 'unary':
		case 'percent':
			return { ...root, operand: replaceNode(root.operand, target, by) };
		case 'binary':
			return {
				...root,
				left: replaceNode(root.left, target, by),
				right: replaceNode(root.right, target, by),
			};
		case 'call':
			return { ...root, args: root.args.map((a) => replaceNode(a, target, by)) };
		case 'invoke':
			return {
				...root,
				callee: replaceNode(root.callee, target, by),
				args: root.args.map((a) => replaceNode(a, target, by)),
			};
		default:
			return root;
	}
}

function literalOf(value: CellValue): FormulaAst {
	if (isCellError(value)) return { type: 'error', code: value.error };
	if (value === null) return { type: 'number', value: 0 };
	if (typeof value === 'number') return { type: 'number', value };
	if (typeof value === 'boolean') return { type: 'boolean', value };
	return { type: 'string', value };
}

function literalFromMatrix(m: CellValue[][]): FormulaAst {
	if (m.length === 1 && m[0]?.length === 1) return literalOf(m[0][0] ?? null);
	return {
		type: 'array',
		rows: m.map((row) =>
			row.map((v) => {
				const lit = literalOf(v);
				return lit.type === 'error'
					? { error: lit.code }
					: lit.type === 'number' || lit.type === 'string' || lit.type === 'boolean'
						? lit.value
						: 0;
			}),
		),
	};
}

export class FormulaStepper {
	#stack: Frame[] = [];
	readonly #start: { sheet: number; row: number; col: number };

	constructor(
		private readonly workbook: Workbook,
		private readonly engine: StepEvaluator,
		start: { sheet: number; row: number; col: number },
	) {
		this.#start = start;
		this.restart();
	}

	/** Whether the cell holds a formula the parser accepts. */
	static canEvaluate(workbook: Workbook, at: { sheet: number; row: number; col: number }): boolean {
		return FormulaStepper.#parse(workbook, at) !== undefined;
	}

	static #parse(
		workbook: Workbook,
		at: { sheet: number; row: number; col: number },
	): FormulaAst | undefined {
		const formula = workbook.sheets[at.sheet]?.rows.get(at.row)?.get(at.col)?.formula;
		if (!formula) return undefined;
		try {
			return parseFormula(formula);
		} catch (e) {
			if (e instanceof FormulaError) return undefined;
			throw e;
		}
	}

	restart(): void {
		const ast = FormulaStepper.#parse(this.workbook, this.#start);
		this.#stack = ast ? [{ ...this.#start, ast }] : [];
	}

	get #top(): Frame | undefined {
		return this.#stack[this.#stack.length - 1];
	}

	/** The cell the next sub-expression references when it is a single formula cell. */
	#stepTarget(): { sheet: number; row: number; col: number } | undefined {
		const top = this.#top;
		const node = top ? nextNode(top.ast) : undefined;
		if (!top || !node || node.type !== 'ref' || !node.ref || node.ref.kind !== 'cell')
			return undefined;
		if (node.prefix && (node.prefix.book !== undefined || node.prefix.sheet2 !== undefined))
			return undefined;
		const name = node.prefix?.sheet;
		const sheet =
			name === undefined
				? top.sheet
				: this.workbook.sheets.findIndex((s) => s.name.toLowerCase() === name.toLowerCase());
		if (sheet < 0) return undefined;
		const at = { sheet, row: node.ref.start.row, col: node.ref.start.col };
		return FormulaStepper.#parse(this.workbook, at) ? at : undefined;
	}

	view(): StepView {
		const top = this.#top;
		if (!top)
			return { reference: '', text: '', finished: true, canStepIn: false, canStepOut: false };
		const node = nextNode(top.ast);
		const printed = printFormula(top.ast, node);
		const sheet = this.workbook.sheets[top.sheet]?.name ?? '';
		const view: StepView = {
			reference: `${quoteSheetName(sheet)}!${formatAddress({ row: top.row, col: top.col })}`,
			text: `=${printed.text}`,
			finished: node === undefined,
			canStepIn: this.#stepTarget() !== undefined,
			canStepOut: this.#stack.length > 1,
		};
		if (printed.span) view.underline = { start: printed.span.start + 1, end: printed.span.end + 1 };
		return view;
	}

	/** Computes the underlined sub-expression and shows its value in its place. */
	evaluate(): void {
		const top = this.#top;
		const node = top ? nextNode(top.ast) : undefined;
		if (!top || !node) return;
		const text = printFormula(node).text;
		let result: FormulaAst;
		try {
			result = literalFromMatrix(this.engine.evaluateArray(text, top));
		} catch {
			result = { type: 'error', code: '#VALUE!' };
		}
		top.ast = replaceNode(top.ast, node, result);
	}

	stepIn(): void {
		const top = this.#top;
		const at = this.#stepTarget();
		const node = top ? nextNode(top.ast) : undefined;
		const ast = at ? FormulaStepper.#parse(this.workbook, at) : undefined;
		if (!at || !ast || !node) return;
		this.#stack.push({ ...at, ast });
	}

	stepOut(): void {
		if (this.#stack.length < 2) return;
		this.#stack.pop();
		this.evaluate();
	}

	/** The value once the formula is fully reduced. */
	result(): CellValue | undefined {
		const top = this.#stack[0];
		if (!top || nextNode(top.ast)) return undefined;
		const a = top.ast;
		if (a.type === 'number' || a.type === 'string' || a.type === 'boolean') return a.value;
		if (a.type === 'error') return { error: a.code };
		return undefined;
	}
}
