// The calc engine: the public API (full and incremental recalculation, invalidation, ad hoc
// evaluation) over the recalculation core in engine-core.ts.
import { type CellRange, cellKey } from '../address';
import { changeAreas } from './change-areas';
import { getCell } from '../cells';
import type { Cell, CellValue, Workbook } from '../model';
import { FormulaError } from './ast';
import { withDateSystem } from './date-serial';
import { GraphPreparation, type PreparedFormulas } from './engine-prepare';
import { EngineStructure } from './engine-structure';
import type { CalcEngine, CalcEngineOptions, CellPosition } from './engine-types';
import { rangeHas, sameRange } from './engine-util';
import { evaluateNode } from './evaluator';
import type { FormulaNode } from './graph';
import { implicitIntersection, toMatrix } from './references';
import { clearFootprint, isSpilledCell } from './spill';
import { type Area, ERR, ErrorSignal, LambdaValue, resetArrayBudget, type Value } from './values';

export {
	type CalcEngine,
	type CalcEngineOptions,
	CIRCULAR_REFERENCE_WARNING,
	type CellPosition,
} from './engine-types';

class Engine extends EngineStructure implements CalcEngine {
	private preparation: GraphPreparation | undefined;
	/** The prepared formulas the first recalculation checks (see `trustsStoredValues`). */
	private prepared: PreparedFormulas | undefined;
	/** Stored spills released when the graph was built: their readers must recalculate. */
	private released: Area[] = [];

	// ---- the graph ----

	prepare(timeRemaining?: () => number): boolean {
		if (this.built || this.needsFull) return true;
		this.preparation ??= this.startPreparation();
		if (!this.preparation.run(timeRemaining)) return false;
		this.finishPreparation(this.preparation);
		return true;
	}

	discardPreparation(): void {
		if (!this.fresh || this.needsFull) return;
		this.preparation = undefined;
		this.prepared = undefined;
		this.structuralSeeds.clear();
		this.nodes.clear();
		this.footprints.clear();
		this.graphChanged();
		this.built = false;
	}

	/** Completes the graph, continuing a preparation that is under way. */
	protected build(): void {
		const preparation = this.preparation ?? this.startPreparation();
		preparation.run();
		this.finishPreparation(preparation);
	}

	private startPreparation(): GraphPreparation {
		this.nodes.clear();
		this.footprints.clear();
		this.footprintIndex = undefined;
		this.prepared = undefined;
		return new GraphPreparation(this.workbook, (s, row, col, cell) =>
			this.createNode(s, row, col, cell, true),
		);
	}

	private finishPreparation(preparation: GraphPreparation): void {
		this.preparation = undefined;
		this.built = true;
		this.adoptStoredSpills(preparation.spills);
		this.graphChanged();
		this.reverseIndex = preparation.reverse;
		this.prepared = preparation.formulas;
	}

	/**
	 * Takes the spills a file stored as their anchors' last footprints, so that overwriting an
	 * anchor or shrinking its result clears the cells it covered before it was ever evaluated here.
	 * Cells marked for an anchor that holds no formula able to spill are released.
	 */
	private adoptStoredSpills(spills: ReadonlyMap<number, ReadonlyMap<number, CellRange>>): void {
		for (const [s, anchors] of spills)
			for (const [key, range] of anchors) {
				const node = this.nodes.get(s)?.get(key);
				if (node && !node.legacy && !node.arrayRange) {
					if (!node.spill && !node.blockedSpill) {
						node.spill = range;
						this.footprints.add(node);
					}
					continue;
				}
				const sheet = this.workbook.sheets[s];
				if (!sheet) continue;
				clearFootprint(sheet, range, range.start.row, range.start.col);
				this.released.push({ sheet: s, range });
			}
	}

	// ---- recalculation ----

	recalculateAll(): void {
		resetArrayBudget();
		this.fresh = false;
		this.prepared = undefined;
		if (this.needsFull) this.releaseAllSpills();
		if (!this.built || this.needsFull) this.build();
		this.needsFull = false;
		this.released = [];
		this.structuralSeeds.clear();
		this.cycles = new Set();
		this.run(new Set(this.allNodes()), true);
	}

	recalculateFrom(changes: CellPosition[]): void {
		resetArrayBudget();
		const seeds = new Set<FormulaNode>();
		if (this.fresh && !this.needsFull) {
			// The first recalculation after opening: like Excel, trust the values the file stored and
			// compute only what the change reaches plus formulas saved without a value. A workbook
			// whose formulas may spill is calculated in full, since spill ranges are only known once
			// their anchors have been evaluated. The changed cells themselves are evaluated below.
			this.fresh = false;
			if (!this.built) this.build();
			if (!this.trustsStoredValues(changes, seeds)) {
				this.recalculateAll();
				return;
			}
		} else if (!this.built || this.needsFull) {
			this.recalculateAll();
			return;
		}
		const changed: Area[] = this.released;
		this.released = [];
		const written: CellPosition[] = [];
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
			written.push(change);
			// A spill or array result overwritten here reruns its anchor; most workbooks have none.
			if (!this.footprints.size) continue;
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
		this.takeStructuralSeeds(seeds);
		const reverse = this.reverse();
		// Formulas reading any changed cell: one query per rectangle of them, not per cell.
		for (const area of [...changed, ...changeAreas(written)])
			reverse.dependents(area.sheet, area.range, seeds);
		for (const node of this.allNodes()) if (node.volatile) seeds.add(node);
		this.run(seeds, false);
	}

	/**
	 * Whether the stored results can stand for a fresh graph: apart from the changed cells, every
	 * formula has legacy or CSE array semantics (none can spill). Collects the formulas stored
	 * without a value into `missing`.
	 */
	protected trustsStoredValues(changes: CellPosition[], missing: Set<FormulaNode>): boolean {
		const prepared = this.prepared;
		this.prepared = undefined;
		if (!prepared) return false;
		const current = (node: FormulaNode): boolean =>
			this.nodes.get(node.sheet)?.get(cellKey(node.row, node.col)) === node;
		// Only a formula that can spill needs the changed cells looked up; most workbooks have none,
		// and a bulk write changes far more cells than there are formulas.
		const dynamic = prepared.dynamic.filter(current);
		if (dynamic.length) {
			const changed = new Map<number, Set<number>>();
			for (const c of changes) {
				let keys = changed.get(c.sheet);
				if (!keys) changed.set(c.sheet, (keys = new Set()));
				keys.add(cellKey(c.row, c.col));
			}
			for (const node of dynamic)
				if (!changed.get(node.sheet)?.has(cellKey(node.row, node.col))) return false;
		}
		for (const node of prepared.unvalued) {
			const cell = this.workbook.sheets[node.sheet]?.rows.get(node.row)?.get(node.col);
			if (current(node) && cell && cell.value === null) missing.add(node);
		}
		return true;
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
		this.structuralSeeds.clear();
		this.releaseAllSpills();
		this.preparation = undefined;
		this.prepared = undefined;
		this.nodes.clear();
		this.footprints.clear();
		this.graphChanged();
		this.built = false;
		this.needsFull = true;
	}

	// ---- ad hoc evaluation ----

	protected evaluateAdHoc(formula: string, at: CellPosition): Value {
		resetArrayBudget();
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
