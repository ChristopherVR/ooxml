// The calc engine: the public API (full and incremental recalculation, invalidation, ad hoc
// evaluation) over the recalculation core in engine-core.ts.
import { cellKey } from '../address.js';
import { getCell } from '../cells.js';
import type { Cell, CellValue, Workbook } from '../model.js';
import { FormulaError } from './ast.js';
import { withDateSystem } from './date-serial.js';
import { EngineCore } from './engine-core.js';
import type { CalcEngine, CalcEngineOptions, CellPosition } from './engine-types.js';
import { rangeHas, sameRange } from './engine-util.js';
import { evaluateNode } from './evaluator.js';
import type { FormulaNode } from './graph.js';
import { implicitIntersection, toMatrix } from './references.js';
import { isSpilledCell } from './spill.js';
import { type Area, ERR, ErrorSignal, LambdaValue, type Value } from './values.js';

export {
	type CalcEngine,
	type CalcEngineOptions,
	CIRCULAR_REFERENCE_WARNING,
	type CellPosition,
} from './engine-types.js';

class Engine extends EngineCore implements CalcEngine {
	// ---- recalculation ----

	recalculateAll(): void {
		if (this.needsFull) this.releaseAllSpills();
		if (!this.built || this.needsFull) this.build();
		this.needsFull = false;
		this.cycles = new Set();
		this.run(new Set(this.allNodes()), true);
	}

	recalculateFrom(changes: CellPosition[]): void {
		if (!this.built || this.needsFull) {
			this.recalculateAll();
			return;
		}
		const seeds = new Set<FormulaNode>();
		const changed: Area[] = [];
		let structural = false;
		for (const change of changes) {
			const sheet = this.workbook.sheets[change.sheet];
			if (!sheet) continue;
			const { row, col } = change;
			const cell = getCell(sheet, row, col);
			if (isSpilledCell(cell) && this.claimedBySpill(change.sheet, cell, row, col))
				delete (cell as { spillAnchor?: unknown }).spillAnchor;
			const existing = this.nodes.get(change.sheet)?.get(cellKey(row, col));
			const formula = cell?.formula;
			const changedKind =
				!!existing &&
				(existing.formula !== formula ||
					!sameRange(existing.arrayRange, cell?.arrayRange) ||
					!!existing.legacy !== (!!cell?.legacyFormula && !cell.arrayRange));
			if (existing && changedKind) {
				this.removeNode(existing, changed);
				structural = true;
			}
			if (formula && (!existing || changedKind)) {
				seeds.add(this.createNode(change.sheet, row, col, cell as Cell, false));
				structural = true;
			} else if (existing && formula) seeds.add(existing);
			changed.push({ sheet: change.sheet, range: { start: { row, col }, end: { row, col } } });
			const anchors: FormulaNode[] = [];
			this.footprintsBy().dependents(
				change.sheet,
				{ start: { row, col }, end: { row, col } },
				anchors,
			);
			for (const anchor of anchors) {
				const fp = anchor.spill ?? anchor.blockedSpill;
				if (fp && rangeHas(fp, row, col) && !(anchor.row === row && anchor.col === col)) {
					seeds.add(anchor);
				}
			}
		}
		if (structural) this.graphChanged();
		const reverse = this.reverse();
		for (const area of changed) reverse.dependents(area.sheet, area.range, seeds);
		for (const node of this.allNodes()) if (node.volatile) seeds.add(node);
		this.run(seeds, false);
	}

	/**
	 * Whether a spilled cell is still part of its anchor's live spill, so a change reported for it
	 * is a user overwrite. A cell whose anchor is blocked (or gone) keeps its marker: it was
	 * restored by undo or is a stale value the anchor may claim again.
	 */
	protected claimedBySpill(
		sheet: number,
		cell: { spillAnchor: { row: number; col: number } },
		row: number,
		col: number,
	): boolean {
		const anchor = this.nodes.get(sheet)?.get(cellKey(cell.spillAnchor.row, cell.spillAnchor.col));
		return !!anchor?.spill && rangeHas(anchor.spill, row, col);
	}

	invalidate(): void {
		this.releaseAllSpills();
		this.nodes.clear();
		this.footprints.clear();
		this.graphChanged();
		this.built = false;
		this.needsFull = true;
	}

	// ---- ad hoc evaluation ----

	protected evaluateAdHoc(formula: string, at: CellPosition): Value {
		const ast = this.parse(formula);
		if (ast instanceof FormulaError) return ERR.NAME;
		const frame = this.frameFor(at.sheet, at.row, at.col, ast);
		return withDateSystem(this.workbook.date1904, () => {
			try {
				return evaluateNode(ast, frame);
			} catch (e) {
				if (e instanceof ErrorSignal) return e.value;
				return ERR.VALUE;
			}
		});
	}

	evaluate(formula: string, at: CellPosition): CellValue {
		const value = this.evaluateAdHoc(formula, at);
		const frame = this.frameFor(at.sheet, at.row, at.col, undefined);
		const scalar = implicitIntersection(value, frame);
		return typeof scalar === 'number' && !Number.isFinite(scalar) ? ERR.NUM : scalar;
	}

	evaluateArray(formula: string, at: CellPosition): CellValue[][] {
		const value = this.evaluateAdHoc(formula, at);
		if (value instanceof LambdaValue) return [[ERR.CALC]];
		try {
			return toMatrix(this, value)
				.block()
				.data.map((row) => [...row]);
		} catch (e) {
			if (e instanceof ErrorSignal) return [[e.value]];
			throw e;
		}
	}

	circularCells(): CellPosition[] {
		return [...this.cycles].map(({ sheet, row, col }) => ({ sheet, row, col }));
	}
}

/** Creates a calculation engine bound to a workbook. */
export function createCalcEngine(workbook: Workbook, options: CalcEngineOptions = {}): CalcEngine {
	return new Engine(workbook, options);
}
