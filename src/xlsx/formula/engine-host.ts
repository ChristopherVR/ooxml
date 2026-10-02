// The calc engine's view of the workbook: formula nodes, cell reads that compute pending
// formulas on demand, and the dependency indexes.
import { cellKey, type CellRange, normalizeRange } from '../address.js';
import type { Cell, Workbook } from '../model.js';
import { FormulaError, type FormulaAst } from './ast.js';
import type { EvalHost } from './context.js';
import type { CalcEngineOptions } from './engine-types.js';
import { analyze, ColumnIndex, type FormulaNode, ReverseIndex } from './graph.js';
import { parseFormula } from './parser.js';
import { scanRange, sheetBounds } from './sheet-scan.js';
import { clearFootprint, isSpilledCell } from './spill.js';
import type { Area, Scalar } from './values.js';

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
	protected footprintChanges: Area[] = [];
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
			if (anchor && pending.has(anchor)) this.compute(anchor);
		}
		if (!cell.formula) return;
		const node = this.nodes.get(sheet)?.get(cellKey(row, col));
		if (!node) return;
		if (this.inProgress.has(node)) {
			this.dynamicCycleHits.add(node);
			this.dynamicCycleHit = true;
			return;
		}
		if (pending.has(node)) this.compute(node);
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
		if (this.pending?.has(node)) this.compute(node);
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
				if (!(e instanceof FormulaError)) throw e;
				ast = e;
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
			volatile: false,
			unsupported: false,
			keepCached: false,
		};
		if (cell.arrayRange) {
			node.arrayRange = normalizeRange(cell.arrayRange);
			this.footprints.add(node);
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
		this.footprints.delete(node);
	}

	protected graphChanged(): void {
		this.columnIndex = undefined;
		this.reverseIndex = undefined;
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
