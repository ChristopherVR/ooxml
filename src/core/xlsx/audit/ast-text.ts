// Formula syntax tree back to text (Evaluate Formula shows partly reduced formulas). Only a
// printer: parsing and evaluation stay with the shared formula parser and engine.
import { type FormulaAst } from '../formula/ast';
import { formatRefSpec } from '../formula/ref-text';

const RANK: Record<string, number> = {
	'=': 1,
	'<>': 1,
	'<': 1,
	'>': 1,
	'<=': 1,
	'>=': 1,
	'&': 2,
	'+': 3,
	'-': 3,
	'*': 4,
	'/': 4,
	'^': 5,
	',': 8,
	' ': 9,
	':': 10,
};
const UNARY = 6;
const PERCENT = 7;

const rankOf = (node: FormulaAst): number =>
	node.type === 'binary' ? (RANK[node.op] ?? 0) : node.type === 'unary' ? UNARY : 99;

export interface PrintedFormula {
	text: string;
	/** Offsets of `mark` in `text` (the leading `=` is not part of `text`). */
	span?: { start: number; end: number };
}

function constantText(v: string | number | boolean | { error: string }): string {
	if (typeof v === 'string') return `"${v.replace(/"/g, '""')}"`;
	if (typeof v === 'object') return v.error;
	if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
	return String(v);
}

/** Prints `ast` without a leading `=`; reports where the node `mark` landed. */
export function printFormula(ast: FormulaAst, mark?: FormulaAst): PrintedFormula {
	let span: PrintedFormula['span'];
	let out = '';
	const emit = (node: FormulaAst): void => {
		const start = out.length;
		write(node);
		if (node === mark) span = { start, end: out.length };
	};
	const child = (node: FormulaAst, minRank: number): void => {
		const wrap = rankOf(node) < minRank || (node.type === 'binary' && node.op === ',');
		if (wrap) out += '(';
		emit(node);
		if (wrap) out += ')';
	};
	const list = (nodes: FormulaAst[]): void => {
		nodes.forEach((n, i) => {
			if (i) out += ',';
			// A union used as an argument keeps its own parentheses so it stays one argument.
			child(n, 0);
		});
	};
	const write = (node: FormulaAst): void => {
		switch (node.type) {
			case 'number':
				out += String(node.value);
				return;
			case 'string':
			case 'boolean':
				out += constantText(node.value);
				return;
			case 'error':
				out += node.code;
				return;
			case 'missing':
				return;
			case 'array':
				out += `{${node.rows.map((row) => row.map(constantText).join(',')).join(';')}}`;
				return;
			case 'ref':
				out += (node.prefix?.text ?? '') + (node.ref ? formatRefSpec(node.ref) : '#REF!');
				if (node.spill) out += '#';
				return;
			case 'name':
				out += (node.prefix?.text ?? '') + node.name;
				return;
			case 'structured':
				out += node.text;
				return;
			case 'unary':
				out += node.op;
				child(node.operand, UNARY);
				return;
			case 'percent':
				child(node.operand, PERCENT);
				out += '%';
				return;
			case 'binary': {
				const rank = RANK[node.op] ?? 0;
				child(node.left, rank);
				out += node.op;
				child(node.right, rank + 1);
				return;
			}
			case 'call':
				out += `${node.rawName}(`;
				list(node.args);
				out += ')';
				return;
			case 'invoke':
				child(node.callee, 99);
				out += '(';
				list(node.args);
				out += ')';
				return;
		}
	};
	emit(ast);
	return span ? { text: out, span } : { text: out };
}
