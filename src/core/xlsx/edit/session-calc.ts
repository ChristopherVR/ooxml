// When an edit session recalculates: incrementally after cell edits and after the structural edits
// the calc engine can follow in place (row and column inserts and deletes, sheet renames), fully
// after other structural ones, and not at all in manual calculation mode until calculateNow /
// calculateSheet.
import type { Workbook } from '../model';
import { calcModeOf } from './calc-mode';
import type { CalcEngine } from './deps';
import { type CalcHint, type HistoryStep, touchedCells } from './history';

type Position = { sheet: number; row: number; col: number };

const isStructural = (step: HistoryStep): boolean =>
	step.structural || step.entries.some((entry) => entry.before.kind !== 'cells');

/** What the engine must follow for a step applied, undone or redone; undefined when it cannot. */
function hintFor(step: HistoryStep, mode: 'apply' | 'undo' | 'redo'): CalcHint | undefined {
	const hint = step.calc;
	if (!hint) return undefined;
	if (hint.kind === 'rename')
		return mode === 'undo' ? { kind: 'rename', from: hint.to, to: hint.from } : hint;
	// Undoing a delete brings back references the graph no longer has: that rebuilds.
	return mode === 'undo' ? undefined : hint;
}

/** The settings a calculation depends on besides the cells (a change recalculates everything). */
const settingsOf = (workbook: Workbook): string =>
	`${workbook.iterate ? `${workbook.iterate.count}:${workbook.iterate.delta}` : ''}|${workbook.date1904}`;

export class SessionCalculator {
	/** Structural edits the engine could not follow leave the dependency graph stale (manual mode). */
	private staleGraph = false;
	/** Cells changed in manual mode since the last calculation. */
	private dirty: Position[] = [];
	/** The settings at the last calculation; undefined before the first. */
	private calculated: string | undefined;

	constructor(
		private readonly workbook: Workbook,
		private readonly calc: CalcEngine,
		/** False when the session was created with `recalc: false`. */
		readonly enabled: boolean,
	) {}

	/** Recalculates after an applied, undone or redone step (failures become warnings). */
	afterStep(step: HistoryStep, mode: 'apply' | 'undo' | 'redo' = 'apply'): void {
		if (!this.enabled) {
			// The engine is not told about this change, so a graph built ahead of time is stale.
			this.calc.discardPreparation();
			return;
		}
		const structural = isStructural(step);
		const hint = structural ? hintFor(step, mode) : undefined;
		if (calcModeOf(this.workbook) === 'manual') {
			// The graph follows a structural edit only while it matches the workbook.
			if (hint && !this.staleGraph && !this.dirty.length) {
				this.guard(() => this.follow(hint));
				return;
			}
			this.calc.discardPreparation();
			if (structural) this.staleGraph = true;
			else this.dirty.push(...touchedCells(step));
			return;
		}
		if (hint && !this.staleGraph) {
			this.guard(() => {
				this.follow(hint);
				this.calc.recalculateFrom(touchedCells(step));
				this.calculated = settingsOf(this.workbook);
			});
			return;
		}
		this.recalc(structural, () => touchedCells(step));
	}

	/** Recalculates after a change applied outside the history (a collaborator's edit). */
	afterExternal(structural: boolean, cells: Position[]): void {
		if (!this.enabled || calcModeOf(this.workbook) === 'manual') {
			this.calc.discardPreparation();
			if (!this.enabled) return;
			if (structural) this.staleGraph = true;
			else this.dirty.push(...cells);
			return;
		}
		this.recalc(structural, () => cells);
	}

	private follow(hint: CalcHint): void {
		if (hint.kind === 'shift') this.calc.shiftCells(hint.sheet, hint.shift);
		else this.calc.renameSheet(hint.from, hint.to);
	}

	private guard(run: () => void): void {
		try {
			run();
		} catch (error) {
			const message = `Recalculation failed: ${error instanceof Error ? error.message : String(error)}`;
			if (!this.workbook.warnings.includes(message)) this.workbook.warnings.push(message);
		}
	}

	/** Builds the dependency graph ahead of the first edit (see `EditSession.prepareCalculation`). */
	prepare(timeRemaining?: () => number): boolean {
		if (!this.enabled || calcModeOf(this.workbook) === 'manual' || this.staleGraph) return true;
		try {
			return this.calc.prepare(timeRemaining);
		} catch {
			// The first edit builds the graph again and reports what failed as a warning.
			this.calc.discardPreparation();
			return true;
		}
	}

	/**
	 * Before a structural edit the graph can follow in place: finishes a graph being prepared (or
	 * builds it) while it still matches the workbook, so the edit moves it instead of rebuilding.
	 */
	beforeStructural(): void {
		if (!this.enabled || calcModeOf(this.workbook) === 'manual' || this.staleGraph) return;
		this.prepare();
	}

	private recalc(structural: boolean, cells: () => Position[]): void {
		this.guard(() => {
			if (this.staleGraph || structural) this.full();
			else {
				const changed = cells();
				if (changed.length) this.calc.recalculateFrom(changed);
				this.calculated = settingsOf(this.workbook);
			}
		});
	}

	/** Rebuilds the dependency graph and recalculates every formula (Ctrl+Alt+F9). */
	full(): void {
		this.staleGraph = false;
		this.dirty = [];
		this.calc.invalidate();
		this.calc.recalculateAll();
		this.calculated = settingsOf(this.workbook);
	}

	/**
	 * Calculate Now (F9): the cells changed since the last calculation, the formulas structural
	 * edits queued, their dependents and the volatile functions. Everything when `full`, before
	 * the session's first calculation, after the iteration settings changed or while the graph
	 * is stale.
	 */
	now(full: boolean): void {
		if (full || this.staleGraph || this.calculated !== settingsOf(this.workbook)) {
			this.full();
			return;
		}
		const changed = this.dirty;
		this.dirty = [];
		this.calc.recalculateFrom(changed);
	}

	/**
	 * Recalculates one sheet's formulas and their dependents (Shift+F9); after structural edits
	 * in manual mode the graph is stale, so the whole workbook is recalculated instead.
	 */
	sheet(s: number): void {
		const sheet = this.workbook.sheets[s];
		if (!sheet) throw new RangeError(`No sheet at index ${s}`);
		if (this.staleGraph) {
			this.full();
			return;
		}
		const cells: Position[] = [];
		for (const [row, line] of sheet.rows)
			for (const [col, cell] of line)
				if (cell.formula !== undefined) cells.push({ sheet: s, row, col });
		if (cells.length) this.calc.recalculateFrom(cells);
	}
}
