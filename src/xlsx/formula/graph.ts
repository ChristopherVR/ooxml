// Static dependency extraction and the indexes the calc engine uses to order and propagate work.
import { cellKey, type CellRange, keyToAddress, rangesIntersect } from '../address.js';
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
	/** Static precedents. */
	deps: Area[];
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
}

interface DepContext {
	host: EvalHost;
	sheet: number;
	row: number;
	col: number;
	deps: Area[];
	volatile: boolean;
	unsupported: boolean;
	depth: number;
}

function visit(node: FormulaAst, c: DepContext): void {
	switch (node.type) {
		case 'ref': {
			if (!node.ref) return;
			const sheets = prefixSheets(c.host, node.prefix, c.sheet);
			if (!sheets) return;
			const range = specRange(node.ref);
			const target = node.spill ? { start: range.start, end: { ...range.start } } : range;
			for (const sheet of sheets) c.deps.push({ sheet, range: target });
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
			visit(node.left, c);
			visit(node.right, c);
			if (node.op === ':' && node.left.type === 'ref' && node.right.type === 'ref') {
				const l = c.deps[c.deps.length - 2];
				const r = c.deps[c.deps.length - 1];
				if (l && r && l.sheet === r.sheet) {
					c.deps.push({
						sheet: l.sheet,
						range: {
							start: {
								row: Math.min(l.range.start.row, r.range.start.row),
								col: Math.min(l.range.start.col, r.range.start.col),
							},
							end: {
								row: Math.max(l.range.end.row, r.range.end.row),
								col: Math.max(l.range.end.col, r.range.end.col),
							},
						},
					});
				}
			}
			return;
		case 'call': {
			const spec = getFunction(node.name);
			if (spec?.volatile) c.volatile = true;
			if (!spec && !isLocalOrDefined(node.rawName, c)) c.unsupported = true;
			for (const arg of node.args) visit(arg, c);
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
	if (value instanceof RefValue) c.deps.push(...value.areas);
}

/** Static precedents of a formula plus whether it is volatile or uses unknown functions. */
export function analyze(
	ast: FormulaAst,
	host: EvalHost,
	sheet: number,
	row: number,
	col: number,
): { deps: Area[]; volatile: boolean; unsupported: boolean } {
	const c: DepContext = {
		host,
		sheet,
		row,
		col,
		deps: [],
		volatile: false,
		unsupported: false,
		depth: 0,
	};
	visit(ast, c);
	return { deps: c.deps, volatile: c.volatile, unsupported: c.unsupported };
}

/** Formula nodes by sheet and column, rows sorted, for "which formulas are inside this range". */
export class ColumnIndex {
	private readonly sheets = new Map<number, Map<number, FormulaNode[]>>();

	constructor(nodes: Iterable<FormulaNode>) {
		for (const node of nodes) {
			let cols = this.sheets.get(node.sheet);
			if (!cols) this.sheets.set(node.sheet, (cols = new Map()));
			let list = cols.get(node.col);
			if (!list) cols.set(node.col, (list = []));
			list.push(node);
		}
		for (const cols of this.sheets.values())
			for (const list of cols.values()) list.sort((a, b) => a.row - b.row);
	}

	nodesIn(area: Area, out: FormulaNode[]): void {
		const cols = this.sheets.get(area.sheet);
		if (!cols) return;
		const { start, end } = area.range;
		const width = end.col - start.col + 1;
		const scan = (list: FormulaNode[]): void => {
			let lo = 0;
			let hi = list.length;
			while (lo < hi) {
				const mid = (lo + hi) >> 1;
				if ((list[mid] as FormulaNode).row < start.row) lo = mid + 1;
				else hi = mid;
			}
			for (let i = lo; i < list.length; i++) {
				const node = list[i] as FormulaNode;
				if (node.row > end.row) break;
				out.push(node);
			}
		};
		if (width <= cols.size) {
			for (let c = start.col; c <= end.col; c++) {
				const list = cols.get(c);
				if (list) scan(list);
			}
		} else {
			for (const [c, list] of cols) if (c >= start.col && c <= end.col) scan(list);
		}
	}
}

const BUCKET_WIDTH = 64;

/** Which formulas read a cell: single-cell deps by key, ranges bucketed by column. */
export class ReverseIndex {
	private readonly cells = new Map<number, Map<number, FormulaNode[]>>();
	private readonly columns = new Map<
		number,
		Map<number, { range: CellRange; node: FormulaNode }[]>
	>();
	private readonly wide = new Map<number, { range: CellRange; node: FormulaNode }[]>();

	constructor(nodes: Iterable<FormulaNode>) {
		for (const node of nodes) for (const dep of node.deps) this.add(dep, node);
	}

	private add(dep: Area, node: FormulaNode): void {
		const { start, end } = dep.range;
		if (start.row === end.row && start.col === end.col) {
			let map = this.cells.get(dep.sheet);
			if (!map) this.cells.set(dep.sheet, (map = new Map()));
			const key = cellKey(start.row, start.col);
			const list = map.get(key);
			if (list) list.push(node);
			else map.set(key, [node]);
			return;
		}
		const entry = { range: dep.range, node };
		if (end.col - start.col + 1 > BUCKET_WIDTH) {
			let list = this.wide.get(dep.sheet);
			if (!list) this.wide.set(dep.sheet, (list = []));
			list.push(entry);
			return;
		}
		let cols = this.columns.get(dep.sheet);
		if (!cols) this.columns.set(dep.sheet, (cols = new Map()));
		for (let c = start.col; c <= end.col; c++) {
			const list = cols.get(c);
			if (list) list.push(entry);
			else cols.set(c, [entry]);
		}
	}

	/** Formulas whose precedents intersect `range` on `sheet`. */
	dependents(sheet: number, range: CellRange, out: Set<FormulaNode>): void {
		const cells = this.cells.get(sheet);
		if (cells) {
			const size = (range.end.row - range.start.row + 1) * (range.end.col - range.start.col + 1);
			if (size <= 4096) {
				for (let r = range.start.row; r <= range.end.row; r++) {
					for (let c = range.start.col; c <= range.end.col; c++) {
						const list = cells.get(cellKey(r, c));
						if (list) for (const node of list) out.add(node);
					}
				}
			} else {
				for (const [key, list] of cells) {
					const { row, col } = keyToAddress(key);
					if (
						row < range.start.row ||
						row > range.end.row ||
						col < range.start.col ||
						col > range.end.col
					)
						continue;
					for (const node of list) out.add(node);
				}
			}
		}
		const cols = this.columns.get(sheet);
		if (cols) {
			const width = range.end.col - range.start.col + 1;
			const check = (list: { range: CellRange; node: FormulaNode }[]): void => {
				for (const entry of list) if (rangesIntersect(entry.range, range)) out.add(entry.node);
			};
			if (width <= cols.size) {
				for (let c = range.start.col; c <= range.end.col; c++) {
					const list = cols.get(c);
					if (list) check(list);
				}
			} else {
				for (const [c, list] of cols) if (c >= range.start.col && c <= range.end.col) check(list);
			}
		}
		for (const entry of this.wide.get(sheet) ?? []) {
			if (rangesIntersect(entry.range, range)) out.add(entry.node);
		}
	}
}
