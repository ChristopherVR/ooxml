// The calc engine's recalculation core: ordering, cycles, spilling and per-formula evaluation.
import { cellKey, type CellRange } from '../address.js';
import type { Cell } from '../model.js';
import type { FormulaAst } from './ast.js';
import type { Frame } from './context.js';
import { withDateSystem } from './date-serial.js';
import { Deferred, EngineHost } from './engine-host.js';
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
			this.evaluatedAt = new Map();
			this.evaluateSet(dirty);
			if (this.footprintChanges.length === 0) break;
			const next = new Set<FormulaNode>();
			const reverse = this.reverse();
			const found: FormulaNode[] = [];
			for (const area of this.footprintChanges) {
				found.length = 0;
				reverse.dependents(area.sheet, area.range, found);
				// A formula evaluated after the spill changed already read its new values.
				for (const node of found)
					if (!((this.evaluatedAt.get(node) ?? -1) > area.at)) next.add(node);
			}
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
		this.blockCache.clear();
		const columns = this.columns();
		const footprints = this.footprintsBy();
		const succCache = new Map<FormulaNode, FormulaNode[]>();
		const successors = (node: FormulaNode): FormulaNode[] => {
			let out = succCache.get(node);
			if (out) return out;
			const found: FormulaNode[] = [];
			for (const dep of node.deps) {
				columns.nodesIn(dep, found);
				const anchors: FormulaNode[] = [];
				footprints.dependents(dep.sheet, dep.range, anchors);
				// A blocked spill does not cover its footprint; only live spills and CSE ranges do.
				for (const anchor of anchors) if (anchor.spill ?? anchor.arrayRange) found.push(anchor);
			}
			out = [...new Set(found)].filter((n) => dirty.has(n));
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
					this.computeFrom(first);
				}
			}
		} finally {
			this.pending = undefined;
			this.blockCache.clear();
		}
	}

	/**
	 * Computes a formula at the top level. Formulas it reads on demand nest inside it; past a depth
	 * limit they unwind with `Deferred` and are computed here first (an explicit stack), so long
	 * INDIRECT chains need no deep recursion. A node needed again while it waits is in a loop.
	 */
	protected computeFrom(root: FormulaNode): void {
		const stack = [root];
		const waiting = new Set(stack);
		while (stack.length) {
			const node = stack[stack.length - 1] as FormulaNode;
			if (!this.pending?.has(node)) {
				waiting.delete(stack.pop() as FormulaNode);
				continue;
			}
			try {
				this.compute(node);
			} catch (e) {
				if (!(e instanceof Deferred)) throw e;
				if (!waiting.has(e.node)) {
					stack.push(e.node);
					waiting.add(e.node);
					continue;
				}
				// Every formula from the needed one up to the top waits on the next: a circular reference.
				for (;;) {
					const member = stack.pop() as FormulaNode;
					waiting.delete(member);
					this.markCycle(member);
					if (member === e.node) break;
				}
			}
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
		} catch (e) {
			// Unwound to compute something it reads first: it is retried, so it is pending again.
			if (e instanceof Deferred) this.pending?.add(node);
			throw e;
		} finally {
			this.inProgress.delete(node);
		}
		if (this.dynamicCycleHits.has(node)) {
			this.cycles.add(node);
			result = 0;
		}
		this.evaluatedAt.set(node, ++this.evaluations);
		this.store(node, cell, result);
	}

	protected store(node: FormulaNode, cell: Cell, result: Scalar | Matrix): void {
		const sheet = this.workbook.sheets[node.sheet];
		if (!sheet) return;
		const before = node.spill;
		const blockedBefore = node.blockedSpill;
		storeResult(sheet, node, cell, result, (anchor, r, c) => {
			const other = this.nodes.get(node.sheet)?.get(cellKey(anchor.row, anchor.col));
			return !!other && other !== node && !!other.spill && rangeHas(other.spill, r, c);
		});
		if (!sameRange(before, node.spill) || !sameRange(blockedBefore, node.blockedSpill)) {
			this.footprintIndex = undefined;
			this.blockCache.clear();
		}
		if (node.spill || node.blockedSpill || node.arrayRange) this.footprints.add(node);
		else this.footprints.delete(node);
		if (!sameRange(before, node.spill)) {
			const ranges = [before, node.spill].filter((r): r is CellRange => r !== undefined);
			this.footprintChanges.push({
				sheet: node.sheet,
				range: boundingBox(ranges),
				at: this.evaluations,
			});
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
		if (!node.ast) return node.parseError ?? ERR.NAME;
		const ast = node.ast;
		const frame = this.frameFor(node.sheet, node.row, node.col, ast, node.legacy);
		return withDateSystem(this.workbook.date1904, () => {
			try {
				const value = evaluateNode(ast, frame);
				return node.legacy ? clean(implicitIntersection(value, frame)) : finalize(value, this);
			} catch (e) {
				if (e instanceof Deferred) throw e;
				if (e instanceof ErrorSignal) return e.value;
				// A formula nested past the stack shows #NUM! instead of escaping the recalculation.
				if (e instanceof RangeError) return ERR.NUM;
				return ERR.VALUE;
			}
		});
	}
}
