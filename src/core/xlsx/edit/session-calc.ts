// When an edit session recalculates: incrementally after cell edits, fully after structural
// ones, and not at all in manual calculation mode until calculateNow / calculateSheet.
import type { Workbook } from '../model.js';
import { calcModeOf } from './calc-mode.js';
import type { CalcEngine } from './deps.js';
import { type HistoryStep, touchedCells } from './history.js';

const isStructural = (step: HistoryStep): boolean =>
	step.structural || step.entries.some((entry) => entry.before.kind !== 'cells');

export class SessionCalculator {
	/** Structural edits made in manual mode leave the dependency graph stale until a full pass. */
	private staleGraph = false;

	constructor(
		private readonly workbook: Workbook,
		private readonly calc: CalcEngine,
		/** False when the session was created with `recalc: false`. */
		readonly enabled: boolean,
	) {}

	/** Recalculates after an applied, undone or redone step (failures become warnings). */
	afterStep(step: HistoryStep): void {
		if (!this.enabled) return;
		if (calcModeOf(this.workbook) === 'manual') {
			if (isStructural(step)) this.staleGraph = true;
			return;
		}
		try {
			if (this.staleGraph || isStructural(step)) this.full();
			else {
				const changed = touchedCells(step);
				if (changed.length) this.calc.recalculateFrom(changed);
			}
		} catch (error) {
			const message = `Recalculation failed: ${error instanceof Error ? error.message : String(error)}`;
			if (!this.workbook.warnings.includes(message)) this.workbook.warnings.push(message);
		}
	}

	/** Rebuilds the dependency graph and recalculates every formula (F9). */
	full(): void {
		this.staleGraph = false;
		this.calc.invalidate();
		this.calc.recalculateAll();
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
		const cells: { sheet: number; row: number; col: number }[] = [];
		for (const [row, line] of sheet.rows)
			for (const [col, cell] of line)
				if (cell.formula !== undefined) cells.push({ sheet: s, row, col });
		if (cells.length) this.calc.recalculateFrom(cells);
	}
}
