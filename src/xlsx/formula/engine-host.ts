// The calc engine's view of the workbook: formula nodes, cell reads that compute pending
// formulas on demand, and the dependency indexes.
import { cellKey, type CellRange, normalizeRange } from '../address.js';
import type { Cell, Workbook } from '../model.js';
import { FormulaError, type FormulaAst } from './ast.js';
import type { EvalHost } from './context.js';
import type { CalcEngineOptions } from './engine-types.js';
import { analyze, type FormulaNode } from './graph.js';
import { AreaIndex, ColumnIndex, ReverseIndex } from './graph-index.js';
import { parseFormula } from './parser.js';
import { scanRange, sheetBounds } from './sheet-scan.js';
import { clearFootprint, isSpilledCell } from './spill.js';
import { type Area, err, Matrix, type Scalar } from './values.js';

/**
 * Thrown to unwind evaluation when formulas computed on demand nest too deeply (a long chain of
 * INDIRECT or OFFSET reads); the engine computes `node` first and then retries, so chains of
 * any length evaluate without growing the call stack.
 */
export class Deferred {
	constructor(readonly node: FormulaNode) {}
}

/** How many formulas may be computed inside one another before the work is deferred. */
const MAX_NESTED_COMPUTE = 40;

export abstract class EngineHost implements EvalHost {
	protected readonly nodes = new Map<number, Map<number, FormulaNode>>();
	protected readonly footprints = new Set<FormulaNode>();
	protected readonly astCache = new Map<string, FormulaAst | FormulaError>();
	protected readonly boundsCache = new Map<number, { rows: number; cols: number }>();
	protected readonly inProgress = new Set<FormulaNode>();
	protected cycles = new Set<FormulaNode>();
	protected dynamicCycleHits = new Set<FormulaNode>();
	protected pending: Set<FormulaNode> | undefined;
	protected columnIndex: ColumnIndex | undefined;
	protected reverseIndex: ReverseIndex | undefined;
	protected footprintIndex: AreaIndex<FormulaNode> | undefined;
	/** Range reads cached for one evaluation pass (see `readBlock`). */
	protected readonly blockCache = new Map<string, Matrix>();
	/** Counts reads that hit a formula still being evaluated (a circular reference). */
	protected cycleReads = 0;
	protected computeDepth = 0;
	/** Spill footprints that changed this round, with the evaluation count when they did. */
	protected footprintChanges: (Area & { at: number })[] = [];
	/** When each formula was last evaluated in this round (an increasing count). */
	protected evaluatedAt = new Map<FormulaNode, number>();
	protected evaluations = 0;
	protected built = false;
	protected needsFull = false;

	constructor(
		readonly workbook: Workbook,
		protected readonly options: CalcEngineOptions,
	) {}

	// ---- EvalHost ----

	readCell(sheet: number, row: number, col: number): Scalar {
		const cell = this.workbook.sheets[sheet]?.rows.get(row)?.get(col);
		if (this.pending && cell) this.ensureComputed(sheet, row, col, cell);
		if (this.dynamicCycleHit) {
			this.dynamicCycleHit = false;
			return 0;
		}
		return this.workbook.sheets[sheet]?.rows.get(row)?.get(col)?.value ?? null;
	}

	protected dynamicCycleHit = false;

	protected ensureComputed(sheet: number, row: number, col: number, cell: Cell): void {
		const pending = this.pending;
		if (!pending) return;
		if (isSpilledCell(cell)) {
			const anchor = this.nodes
				.get(sheet)
				?.get(cellKey(cell.spillAnchor.row, cell.spillAnchor.col));
			if (anchor && pending.has(anchor)) this.computeNested(anchor);
		}
		if (!cell.formula) return;
		const node = this.nodes.get(sheet)?.get(cellKey(row, col));
		if (!node) return;
		if (this.inProgress.has(node)) {
			this.dynamicCycleHits.add(node);
			this.dynamicCycleHit = true;
			this.cycleReads++;
			return;
		}
		if (pending.has(node)) this.computeNested(node);
	}

	/** Computes a pending formula read by the one being evaluated, deferring when nested deeply. */
	protected computeNested(node: FormulaNode): void {
		if (this.computeDepth >= MAX_NESTED_COMPUTE) throw new Deferred(node);
		this.computeDepth++;
		try {
			this.compute(node);
		} finally {
			this.computeDepth--;
		}
	}

	/**
	 * A range's values, cached until the evaluation pass ends or a spill changes shape. Pending
	 * formulas inside are computed while reading, so later readers (SUMIF over the same column
	 * in thousands of cells) reuse the matrix instead of reading every cell again.
	 */
	readBlock(sheet: number, range: CellRange): Matrix {
		const { start, end } = range;
		const key = `${sheet}:${start.row}:${start.col}:${end.row}:${end.col}`;
		const cached = this.pending ? this.blockCache.get(key) : undefined;
		if (cached) return cached;
		const before = this.cycleReads;
		const block = Matrix.build(end.row - start.row + 1, end.col - start.col + 1, (r, c) =>
			this.readCell(sheet, start.row + r, start.col + c),
		);
		if (this.pending && this.cycleReads === before) this.blockCache.set(key, block);
		return block;
	}

	forEachStored(
		sheet: number,
		range: CellRange,
		visit: (value: Scalar, row: number, col: number) => void,
	): void {
		const ws = this.workbook.sheets[sheet];
		if (!ws) return;
		scanRange(ws, range, (cell, row, col) => {
			if (this.pending) this.ensureComputed(sheet, row, col, cell);
			const value = this.dynamicCycleHit ? 0 : (ws.rows.get(row)?.get(col)?.value ?? null);
			this.dynamicCycleHit = false;
			if (value !== null) visit(value, row, col);
		});
	}

	bounds(sheet: number): { rows: number; cols: number } {
		let cached = this.boundsCache.get(sheet);
		if (!cached) {
			const ws = this.workbook.sheets[sheet];
			cached = ws ? sheetBounds(ws) : { rows: 0, cols: 0 };
			this.boundsCache.set(sheet, cached);
		}
		return cached;
	}

	spillRange(sheet: number, row: number, col: number): CellRange | undefined {
		const node = this.nodes.get(sheet)?.get(cellKey(row, col));
		if (!node) return undefined;
		if (this.pending?.has(node)) this.computeNested(node);
		return node.spill ? normalizeRange(node.spill) : undefined;
	}

	cellFormula(sheet: number, row: number, col: number): string | undefined {
		return this.workbook.sheets[sheet]?.rows.get(row)?.get(col)?.formula;
	}

	parse(formula: string): FormulaAst | FormulaError {
		let ast = this.astCache.get(formula);
		if (!ast) {
			try {
				ast = parseFormula(formula);
			} catch (e) {
				// A stack overflow while parsing is reported like a too-deeply nested formula.
				if (e instanceof RangeError) ast = new FormulaError(e.message, -1, '#VALUE!');
				else if (!(e instanceof FormulaError)) throw e;
				else ast = e;
			}
			if (this.astCache.size > 100_000) this.astCache.clear();
			this.astCache.set(formula, ast);
		}
		return ast;
	}

	now(): Date {
		return this.options.now ? this.options.now() : new Date();
	}

	random(): number {
		return this.options.random ? this.options.random() : Math.random();
	}

	// ---- graph ----

	protected build(): void {
		this.nodes.clear();
		this.footprints.clear();
		this.footprintIndex = undefined;
		this.workbook.sheets.forEach((sheet, s) => {
			for (const [row, cells] of sheet.rows) {
				for (const [col, cell] of cells) {
					if (cell.formula) this.createNode(s, row, col, cell, true);
				}
			}
		});
		this.built = true;
		this.graphChanged();
	}

	protected createNode(
		sheet: number,
		row: number,
		col: number,
		cell: Cell,
		initial: boolean,
	): FormulaNode {
		const formula = cell.formula ?? '';
		const parsed = this.parse(formula);
		const node: FormulaNode = {
			sheet,
			row,
			col,
			formula,
			ast: parsed instanceof FormulaError ? undefined : parsed,
			deps: [],
			dynamicDeps: [],
			volatile: false,
			unsupported: false,
			keepCached: false,
		};
		if (parsed instanceof FormulaError && parsed.code) node.parseError = err(parsed.code);
		if (cell.arrayRange) {
			node.arrayRange = normalizeRange(cell.arrayRange);
			this.footprints.add(node);
			this.footprintIndex = undefined;
		} else if (cell.legacyFormula) node.legacy = true;
		if (node.ast) Object.assign(node, analyze(node.ast, this, sheet, row, col));
		node.keepCached = initial && (!node.ast || node.unsupported) && cell.value !== null;
		let map = this.nodes.get(sheet);
		if (!map) this.nodes.set(sheet, (map = new Map()));
		map.set(cellKey(row, col), node);
		return node;
	}

	protected removeNode(node: FormulaNode, changed: Area[]): void {
		const sheet = this.workbook.sheets[node.sheet];
		if (node.spill && sheet) {
			clearFootprint(sheet, node.spill, node.row, node.col);
			changed.push({ sheet: node.sheet, range: node.spill });
		}
		this.nodes.get(node.sheet)?.delete(cellKey(node.row, node.col));
		if (this.footprints.delete(node)) this.footprintIndex = undefined;
	}

	protected graphChanged(): void {
		this.columnIndex = undefined;
		this.reverseIndex = undefined;
		this.footprintIndex = undefined;
	}

	/** Formulas by the area their spill or CSE array covers (spatial lookup for every read). */
	protected footprintsBy(): AreaIndex<FormulaNode> {
		let index = this.footprintIndex;
		if (!index) {
			index = new AreaIndex<FormulaNode>();
			for (const node of this.footprints) {
				const fp = node.spill ?? node.blockedSpill ?? node.arrayRange;
				if (fp) index.add({ sheet: node.sheet, range: fp }, node);
			}
			this.footprintIndex = index;
		}
		return index;
	}

	protected allNodes(): FormulaNode[] {
		const out: FormulaNode[] = [];
		for (const map of this.nodes.values()) for (const node of map.values()) out.push(node);
		return out;
	}

	protected columns(): ColumnIndex {
		return (this.columnIndex ??= new ColumnIndex(this.allNodes()));
	}

	protected reverse(): ReverseIndex {
		return (this.reverseIndex ??= new ReverseIndex(this.allNodes()));
	}

	/** Evaluates a pending formula node now (defined by the engine). */
	protected abstract compute(node: FormulaNode): void;
}
