// Structural edits without a rebuild: a row or column insert or delete moves the formula graph in
// place, and a sheet rename only refreshes the formulas whose text changed. Formulas whose value
// can change are queued for the next recalculation; everything else keeps its value.
import { cellKey, normalizeRange } from '../address';
import { FormulaError } from './ast';
import { EngineCore } from './engine-core';
import { sameRange } from './engine-util';
import { analyze, type FormulaNode } from './graph';
import {
	areaDelta,
	astTraits,
	type AstTraits,
	type GraphShift,
	intersectionMoves,
	intervalDelta,
	movedRange,
	reachesEdit,
} from './shift-graph';
import { isSpilledCell } from './spill';
import { type Area, err } from './values';

const sameAreas = (a: readonly Area[], b: readonly Area[]): boolean =>
	a.length === b.length &&
	a.every((x, i) => {
		const y = b[i] as Area;
		return x.sheet === y.sheet && sameRange(x.range, y.range);
	});

export abstract class EngineStructure extends EngineCore {
	/** Formulas a structural edit left to recalculate (with their dependents) on the next pass. */
	protected readonly structuralSeeds = new Set<FormulaNode>();

	/** Forgets the graph so the next recalculation rebuilds it (defined by the engine). */
	abstract invalidate(): void;

	/** Whether the graph describes the workbook as it was before the edit just reported. */
	private graphReady(): boolean {
		return this.built && !this.needsFull;
	}

	/** Parses a node's current text again and recomputes its precedents. */
	protected reanalyze(node: FormulaNode): void {
		const parsed = this.parse(node.formula);
		delete node.stale;
		delete node.parseError;
		if (parsed instanceof FormulaError) {
			node.ast = undefined;
			if (parsed.code) node.parseError = err(parsed.code);
			node.deps = [];
			node.dynamicDeps = [];
			return;
		}
		node.ast = parsed;
		const found = analyze(parsed, this, node.sheet, node.row, node.col);
		node.deps = found.deps;
		node.dynamicDeps = found.dynamicDeps;
		node.volatile = found.volatile;
		node.unsupported = found.unsupported;
	}

	/**
	 * Moves the graph for rows or columns inserted or deleted on `sheet`, after the workbook was
	 * changed (cells moved, formulas rewritten). Spills on the sheet, array formulas the edit cuts
	 * and anything unexpected fall back to a rebuild on the next recalculation.
	 */
	shiftCells(sheet: number, shift: GraphShift): void {
		if (!this.graphReady() || !this.shiftable(sheet, shift)) {
			this.invalidate();
			return;
		}
		const ws = this.workbook.sheets[sheet];
		const own = this.nodes.get(sheet);
		if (!ws) return;
		// Rekey the edited sheet's formulas first, so the rest of the pass reads the new positions.
		const moved = new Map<FormulaNode, number>();
		if (own) {
			const next = new Map<number, FormulaNode>();
			for (const node of own.values()) {
				const pos = shift.axis === 'row' ? node.row : node.col;
				const delta = intervalDelta(pos, pos, shift);
				if (delta === undefined) {
					this.dropNode(node);
					continue;
				}
				if (delta !== 0) {
					moved.set(node, delta);
					if (shift.axis === 'row') node.row += delta;
					else node.col += delta;
				}
				next.set(cellKey(node.row, node.col), node);
			}
			this.nodes.set(sheet, next);
		}
		for (const map of this.nodes.values())
			for (const node of map.values()) {
				if (!this.shiftNode(node, sheet, shift, moved.get(node) ?? 0)) {
					this.invalidate();
					return;
				}
			}
		this.boundsCache.clear();
		this.graphChanged();
	}

	/** No spill or stale spilled value on the sheet, and no array formula the edit cuts. */
	private shiftable(sheet: number, shift: GraphShift): boolean {
		for (const node of this.footprints) {
			if (node.sheet !== sheet) continue;
			if (node.spill || node.blockedSpill) return false;
			if (
				node.arrayRange &&
				areaDelta({ sheet, range: node.arrayRange }, sheet, shift) === undefined
			)
				return false;
		}
		const ws = this.workbook.sheets[sheet];
		let spilled = false;
		ws?.rows.forEach((cells) => {
			if (!spilled) cells.forEach((cell) => (spilled ||= isSpilledCell(cell)));
		});
		return !spilled;
	}

	private dropNode(node: FormulaNode): void {
		this.structuralSeeds.delete(node);
		this.cycles.delete(node);
		this.footprints.delete(node);
	}

	/** Updates one formula for the edit; false when the graph cannot follow it. */
	private shiftNode(
		node: FormulaNode,
		sheet: number,
		shift: GraphShift,
		nodeDelta: number,
	): boolean {
		const cell = this.workbook.sheets[node.sheet]?.rows.get(node.row)?.get(node.col);
		if (cell?.formula === undefined) return false;
		if (node.arrayRange) {
			const delta = areaDelta({ sheet: node.sheet, range: node.arrayRange }, sheet, shift);
			if (delta === undefined) return false;
			node.arrayRange = movedRange(node.arrayRange, shift.axis, delta);
			if (!cell.arrayRange || !sameRange(node.arrayRange, normalizeRange(cell.arrayRange)))
				return false;
		}
		const changed = cell.formula !== node.formula;
		const traits: AstTraits | undefined = node.ast && astTraits(node.ast);
		const reached =
			reachesEdit(node.deps, sheet, shift) || reachesEdit(node.dynamicDeps, sheet, shift);
		// Most formulas neither move nor read anything that does.
		if (!changed && !reached && nodeDelta === 0 && !traits?.descriptive && !traits?.opaque)
			return true;
		if (traits?.complex) {
			node.formula = cell.formula;
			this.reanalyze(node);
			this.structuralSeeds.add(node);
			return true;
		}
		const pos = shift.axis === 'row' ? node.row - nodeDelta : node.col - nodeDelta;
		const moves = { seed: !!traits?.descriptive, moved: false, cut: false };
		const deps = this.moveAreas(node.deps, sheet, shift, pos, nodeDelta, moves);
		const dynamicDeps = this.moveAreas(node.dynamicDeps, sheet, shift, pos, nodeDelta, moves);
		const { cut, moved: anyMoved } = moves;
		let seed = moves.seed;
		// The rewritten text must agree with the moved precedents; otherwise read it again.
		if (cut || changed !== anyMoved) {
			node.formula = cell.formula;
			this.reanalyze(node);
			this.structuralSeeds.add(node);
			return true;
		}
		if (traits?.positional && (anyMoved || nodeDelta !== 0)) seed = true;
		node.deps = deps;
		node.dynamicDeps = dynamicDeps;
		if (changed) {
			node.formula = cell.formula;
			node.stale = true;
		}
		if (seed) this.structuralSeeds.add(node);
		return true;
	}

	/**
	 * Areas moved for the edit (the same list when none moves). Records in `moves` whether one
	 * moved, one was cut (deleted, widened, pushed off the grid) and whether implicit intersection
	 * with one can now pick another cell.
	 */
	private moveAreas(
		list: Area[],
		sheet: number,
		shift: GraphShift,
		pos: number,
		nodeDelta: number,
		moves: { seed: boolean; moved: boolean; cut: boolean },
	): Area[] {
		let out: Area[] | undefined;
		for (let i = 0; i < list.length; i++) {
			const area = list[i] as Area;
			const delta = areaDelta(area, sheet, shift);
			if (delta === undefined) {
				moves.cut = true;
				continue;
			}
			if (intersectionMoves(area, delta, pos, nodeDelta, shift.axis)) moves.seed = true;
			if (delta === 0) continue;
			moves.moved = true;
			out ??= list.slice();
			out[i] = { sheet: area.sheet, range: movedRange(area.range, shift.axis, delta) };
		}
		return out ?? list;
	}

	/**
	 * Follows a sheet rename (`from` to `to`, formulas already rewritten): precedents are kept by
	 * sheet index, so only rewritten formulas need their text parsed again, and only formulas that
	 * name the new sheet without having been rewritten (previously unresolved), describe sheets
	 * (CELL, SHEET, FORMULATEXT) or read defined names whose areas changed are recalculated.
	 */
	renameSheet(from: string, to: string): void {
		if (!this.graphReady()) {
			this.invalidate();
			return;
		}
		const target = to.toLowerCase();
		let depsChanged = false;
		for (const map of this.nodes.values())
			for (const node of map.values()) {
				const cell = this.workbook.sheets[node.sheet]?.rows.get(node.row)?.get(node.col);
				if (cell?.formula === undefined) {
					this.invalidate();
					return;
				}
				const changed = cell.formula !== node.formula;
				const traits = node.ast && astTraits(node.ast);
				node.formula = cell.formula;
				if (traits?.complex || (!changed && traits?.sheets.has(target))) {
					const deps = node.deps;
					const dynamicDeps = node.dynamicDeps;
					this.reanalyze(node);
					const same = sameAreas(deps, node.deps) && sameAreas(dynamicDeps, node.dynamicDeps);
					if (!same) depsChanged = true;
					if (!same || !traits?.complex) this.structuralSeeds.add(node);
				} else if (changed) node.stale = true;
				if (traits?.descriptive || traits?.opaque) this.structuralSeeds.add(node);
			}
		if (depsChanged) this.graphChanged();
	}

	/** Takes the formulas queued by structural edits. */
	protected takeStructuralSeeds(into: Set<FormulaNode>): void {
		for (const node of this.structuralSeeds)
			if (this.nodes.get(node.sheet)?.get(cellKey(node.row, node.col)) === node) into.add(node);
		this.structuralSeeds.clear();
	}
}
