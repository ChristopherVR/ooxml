// Static dependency extraction: which areas each formula reads (graph-index.ts indexes them).
import type { CellRange } from '../address.js';
import type { CellError } from '../model.js';
import { FormulaError, type FormulaAst } from './ast.js';
import type { EvalHost } from './context.js';
import { findDefinedName } from './evaluator.js';
import { getFunction } from './functions/registry.js';
import { prefixSheets, resolveStructured, specRange } from './references.js';
import { type Area, RefValue } from './values.js';

/** A formula cell known to the engine. */
export interface FormulaNode {
	sheet: number;
	row: number;
	col: number;
	formula: string;
	ast: FormulaAst | undefined;
	/** Static precedents: evaluation order and cycle detection follow these. */
	deps: Area[];
	/**
	 * Areas read only through reference arguments (OFFSET's base, INDEX's array, ROW's
	 * reference) or spanned by `:` between computed references. They only drive incremental
	 * recalculation: the cells actually read are computed on demand, so `=OFFSET(B2,-1,0)+1` in
	 * B2 is not a circular reference (a real loop is still found while evaluating).
	 */
	dynamicDeps: Area[];
	volatile: boolean;
	/** Uses a function this engine does not implement. */
	unsupported: boolean;
	/** Keep the value loaded from the file (unsupported or unparsable formulas). */
	keepCached: boolean;
	/** CSE array formula range. */
	arrayRange?: CellRange;
	/** Pre-dynamic-array semantics (implicit intersection, no spilling); see `Cell.legacyFormula`. */
	legacy?: boolean;
	/** Current dynamic-array spill footprint (anchor included). */
	spill?: CellRange;
	/** Where the formula would spill if it were not blocked. */
	blockedSpill?: CellRange;
	/** The error shown for a formula that cannot be parsed for a reason other than syntax. */
	parseError?: CellError;
}

interface DepContext {
	host: EvalHost;
	sheet: number;
	row: number;
	col: number;
	deps: Area[];
	dynamicDeps: Area[];
	/** Inside a reference argument whose cells are not read as a whole. */
	dynamic: number;
	volatile: boolean;
	unsupported: boolean;
	depth: number;
}

/**
 * Functions whose leading reference argument is not read as a whole: OFFSET and INDEX pick
 * cells from it, ROW / COLUMN / ROWS / COLUMNS / AREAS / ISREF only look at its shape.
 */
const REFERENCE_ARGUMENT = new Set([
	'OFFSET',
	'INDEX',
	'ROW',
	'COLUMN',
	'ROWS',
	'COLUMNS',
	'AREAS',
	'ISREF',
]);

const push = (c: DepContext, area: Area): void => {
	(c.dynamic > 0 ? c.dynamicDeps : c.deps).push(area);
};

/** The bounding box, per sheet, of areas (what `:` between computed references can span). */
function spans(areas: readonly Area[]): Area[] {
	const boxes = new Map<number, Area>();
	for (const { sheet, range } of areas) {
		const box = boxes.get(sheet);
		if (!box) {
			boxes.set(sheet, { sheet, range: { start: { ...range.start }, end: { ...range.end } } });
			continue;
		}
		box.range.start.row = Math.min(box.range.start.row, range.start.row);
		box.range.start.col = Math.min(box.range.start.col, range.start.col);
		box.range.end.row = Math.max(box.range.end.row, range.end.row);
		box.range.end.col = Math.max(box.range.end.col, range.end.col);
	}
	return [...boxes.values()];
}

function visit(node: FormulaAst, c: DepContext): void {
	switch (node.type) {
		case 'ref': {
			if (!node.ref) return;
			const sheets = prefixSheets(c.host, node.prefix, c.sheet);
			if (!sheets) return;
			const range = specRange(node.ref);
			const target = node.spill ? { start: range.start, end: { ...range.start } } : range;
			for (const sheet of sheets) push(c, { sheet, range: target });
			return;
		}
		case 'name': {
			let sheet = c.sheet;
			if (node.prefix) {
				const sheets = prefixSheets(c.host, node.prefix, c.sheet);
				const first = sheets?.[0];
				if (first === undefined) return;
				sheet = first;
			}
			const defined = findDefinedName(c.host, node.name, sheet, node.prefix !== undefined);
			if (defined) {
				if (c.depth > 16) return;
				const ast = c.host.parse(defined.formula);
				if (ast instanceof FormulaError) return;
				c.depth++;
				visit(ast, c);
				c.depth--;
				return;
			}
			addStructured({ table: node.name, specials: [] }, c);
			return;
		}
		case 'structured':
			addStructured(node.ref, c);
			return;
		case 'unary':
		case 'percent':
			visit(node.operand, c);
			return;
		case 'binary':
			visitBinary(node, c);
			return;
		case 'call': {
			const spec = getFunction(node.name);
			if (spec?.volatile) c.volatile = true;
			if (!spec && !isLocalOrDefined(node.rawName, c)) c.unsupported = true;
			const reference = spec !== undefined && REFERENCE_ARGUMENT.has(node.name);
			node.args.forEach((arg, i) => {
				if (reference && i === 0) {
					c.dynamic++;
					visit(arg, c);
					c.dynamic--;
				} else visit(arg, c);
			});
			return;
		}
		case 'invoke':
			visit(node.callee, c);
			for (const arg of node.args) visit(arg, c);
			return;
		default:
			return;
	}
}

/**
 * A binary chain, walking the left spine so long chains (1+1+...+1) do not recurse per term.
 * `A1:B5` depends on the whole box; `:` between computed references (names, INDEX, CHOOSE,
 * OFFSET...) depends dynamically on the box of everything its operands mention.
 */
function visitBinary(node: Extract<FormulaAst, { type: 'binary' }>, c: DepContext): void {
	const spine: Extract<FormulaAst, { type: 'binary' }>[] = [];
	let leftmost: FormulaAst = node;
	while (leftmost.type === 'binary') {
		spine.push(leftmost);
		leftmost = leftmost.left;
	}
	const staticStart = c.deps.length;
	const dynamicStart = c.dynamicDeps.length;
	visit(leftmost, c);
	for (let i = spine.length - 1; i >= 0; i--) {
		const op = spine[i] as (typeof spine)[number];
		visit(op.right, c);
		if (op.op !== ':') continue;
		if (op.left.type === 'ref' && op.right.type === 'ref') {
			const list = c.dynamic > 0 ? c.dynamicDeps : c.deps;
			const l = list[list.length - 2];
			const r = list[list.length - 1];
			if (l && r && l.sheet === r.sheet) push(c, spans([l, r])[0] as Area);
			continue;
		}
		const mentioned = [...c.deps.slice(staticStart), ...c.dynamicDeps.slice(dynamicStart)];
		c.dynamicDeps.push(...spans(mentioned));
	}
}

function isLocalOrDefined(name: string, c: DepContext): boolean {
	// LET / LAMBDA variables are not known statically; treat any defined name as a possible lambda.
	return findDefinedName(c.host, name, c.sheet, false) !== undefined || /^_xlpm\./i.test(name);
}

function addStructured(ref: Parameters<typeof resolveStructured>[0], c: DepContext): void {
	const value = resolveStructured(ref, {
		host: c.host,
		sheet: c.sheet,
		row: c.row,
		col: c.col,
		scope: undefined,
		depth: 0,
	});
	if (value instanceof RefValue) for (const area of value.areas) push(c, area);
}

/** Static precedents of a formula plus whether it is volatile or uses unknown functions. */
export function analyze(
	ast: FormulaAst,
	host: EvalHost,
	sheet: number,
	row: number,
	col: number,
): { deps: Area[]; dynamicDeps: Area[]; volatile: boolean; unsupported: boolean } {
	const c: DepContext = {
		host,
		sheet,
		row,
		col,
		deps: [],
		dynamicDeps: [],
		dynamic: 0,
		volatile: false,
		unsupported: false,
		depth: 0,
	};
	visit(ast, c);
	return {
		deps: c.deps,
		dynamicDeps: c.dynamicDeps,
		volatile: c.volatile,
		unsupported: c.unsupported,
	};
}
