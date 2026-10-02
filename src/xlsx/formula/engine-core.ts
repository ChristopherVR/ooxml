// The calc engine's recalculation core: ordering, cycles, spilling and per-formula evaluation.
import { cellKey, type CellRange, rangesIntersect } from '../address.js';
import type { Cell } from '../model.js';
import type { FormulaAst } from './ast.js';
import type { Frame } from './context.js';
import { withDateSystem } from './date-serial.js';
import { EngineHost } from './engine-host.js';
import { CIRCULAR_REFERENCE_WARNING } from './engine-types.js';
import { boundingBox, clean, finalize, rangeHas, sameRange } from './engine-util.js';
import { evaluateNode } from './evaluator.js';
import type { FormulaNode } from './graph.js';
import { implicitIntersection } from './references.js';
import { stronglyConnected } from './scc.js';
import { isSpilledCell, releaseSpilledCell } from './spill.js';
import { storeResult } from './store.js';
import { ERR, ErrorSignal, type Matrix, type Scalar } from './values.js';

const MAX_ROUNDS = 8;

export abstract class EngineCore extends EngineHost {
	protected releaseAllSpills(): void {
		for (const sheet of this.workbook.sheets) {
			for (const [row, cells] of [...sheet.rows]) {
				for (const [col, cell] of [...cells])
					if (isSpilledCell(cell)) releaseSpilledCell(sheet, row, col);
			}
		}
	}

	protected run(seeds: Set<FormulaNode>, everything: boolean): void {
		let current = seeds;
		for (let round = 0; round < MAX_ROUNDS && current.size; round++) {
			const dirty = everything && round === 0 ? current : this.closure(current);
			this.footprintChanges = [];
			this.evaluateSet(dirty);
			if (this.footprintChanges.length === 0) break;
			const next = new Set<FormulaNode>();
			const reverse = this.reverse();
			for (const area of this.footprintChanges) reverse.dependents(area.sheet, area.range, next);
			current = next;
		}
		if (this.cycles.size && !this.workbook.warnings.includes(CIRCULAR_REFERENCE_WARNING)) {
			this.workbook.warnings.push(CIRCULAR_REFERENCE_WARNING);
		}
	}

	/** Seeds plus every formula that (transitively) reads their cells or spill ranges. */
	protected closure(seeds: Set<FormulaNode>): Set<FormulaNode> {
		const out = new Set(seeds);
		const queue = [...seeds];
		const reverse = this.reverse();
		const found = new Set<FormulaNode>();
		while (queue.length) {
			const node = queue.pop() as FormulaNode;
			found.clear();
			reverse.dependents(
				node.sheet,
				{ start: { row: node.row, col: node.col }, end: { row: node.row, col: node.col } },
				found,
			);
			const fp = node.spill ?? node.arrayRange;
			if (fp) reverse.dependents(node.sheet, fp, found);
			for (const dep of found) {
				if (!out.has(dep)) {
					out.add(dep);
					queue.push(dep);
				}
			}
		}
		return out;
	}

	protected evaluateSet(dirty: Set<FormulaNode>): void {
		this.boundsCache.clear();
		const columns = this.columns();
		const succCache = new Map<FormulaNode, FormulaNode[]>();
		const successors = (node: FormulaNode): FormulaNode[] => {
			let out = succCache.get(node);
			if (out) return out;
			const found: FormulaNode[] = [];
			for (const dep of node.deps) {
				columns.nodesIn(dep, found);
				for (const anchor of this.footprints) {
					const fp = anchor.spill ?? anchor.arrayRange;
					if (fp && anchor.sheet === dep.sheet && rangesIntersect(fp, dep.range))
						found.push(anchor);
				}
			}
			out = found.filter((n) => dirty.has(n));
			succCache.set(node, out);
			return out;
		};
		const order = stronglyConnected(dirty, successors);
		this.pending = new Set(dirty);
		this.dynamicCycleHits = new Set();
		try {
			for (const component of order) {
				const first = component[0] as FormulaNode;
				if (component.length > 1 || successors(first).includes(first)) {
					for (const node of component) this.markCycle(node);
				} else if (this.pending.has(first)) {
					this.compute(first);
				}
			}
		} finally {
			this.pending = undefined;
		}
	}

	protected markCycle(node: FormulaNode): void {
		this.pending?.delete(node);
		this.cycles.add(node);
		const cell = this.workbook.sheets[node.sheet]?.rows.get(node.row)?.get(node.col);
		if (cell && !node.keepCached) this.store(node, cell, 0);
	}

	protected compute(node: FormulaNode): void {
		this.pending?.delete(node);
		const cell = this.workbook.sheets[node.sheet]?.rows.get(node.row)?.get(node.col);
		if (!cell || cell.formula !== node.formula || node.keepCached) return;
		this.inProgress.add(node);
		let result: Scalar | Matrix;
		try {
			result = this.evaluateFormula(node);
		} finally {
			this.inProgress.delete(node);
		}
		if (this.dynamicCycleHits.has(node)) {
			this.cycles.add(node);
			result = 0;
		}
		this.store(node, cell, result);
	}

	protected store(node: FormulaNode, cell: Cell, result: Scalar | Matrix): void {
		const sheet = this.workbook.sheets[node.sheet];
		if (!sheet) return;
		const before = node.spill;
		storeResult(sheet, node, cell, result, (anchor, r, c) => {
			const other = this.nodes.get(node.sheet)?.get(cellKey(anchor.row, anchor.col));
			return !!other && other !== node && !!other.spill && rangeHas(other.spill, r, c);
		});
		if (node.spill || node.blockedSpill || node.arrayRange) this.footprints.add(node);
		else this.footprints.delete(node);
		if (!sameRange(before, node.spill)) {
			const ranges = [before, node.spill].filter((r): r is CellRange => r !== undefined);
			this.footprintChanges.push({ sheet: node.sheet, range: boundingBox(ranges) });
			const bounds = this.boundsCache.get(node.sheet);
			if (bounds && node.spill) {
				bounds.rows = Math.max(bounds.rows, node.spill.end.row + 1);
				bounds.cols = Math.max(bounds.cols, node.spill.end.col + 1);
			}
		}
	}

	protected frameFor(
		sheet: number,
		row: number,
		col: number,
		root: FormulaAst | undefined,
		legacy = false,
	): Frame {
		const frame: Frame = { host: this, sheet, row, col, scope: undefined, depth: 0 };
		if (legacy) Object.assign(frame, { legacy: true });
		return root ? { ...frame, root } : frame;
	}

	protected evaluateFormula(node: FormulaNode): Scalar | Matrix {
		if (!node.ast) return ERR.NAME;
		const ast = node.ast;
		const frame = this.frameFor(node.sheet, node.row, node.col, ast, node.legacy);
		return withDateSystem(this.workbook.date1904, () => {
			try {
				const value = evaluateNode(ast, frame);
				return node.legacy ? clean(implicitIntersection(value, frame)) : finalize(value, this);
			} catch (e) {
				if (e instanceof ErrorSignal) return e.value;
				if (e instanceof RangeError) return ERR.NUM;
				return ERR.VALUE;
			}
		});
	}
}
