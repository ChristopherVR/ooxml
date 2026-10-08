// Changes that did not start in this session: a collaborator's edit or a shared undo applied to
// the workbook by a binding (see `ooxml-core/xlsx/collab`). They bypass the undo history, are
// recalculated like a local step and reach `onChange` listeners flagged `external`.
import type { CellRange } from '../address';
import type { SessionCalculator } from './session-calc';
import type { WorkbookChange, WorkbookChangeKind } from './types';

/** What an external change touched, returned by the function passed to `applyExternal`. */
export interface ExternalChange {
	label: string;
	/** Default `remote`; a shared undo or redo reports `undo` or `redo`. */
	kind?: WorkbookChangeKind;
	/** Rows, columns or sheets changed: views rebuild metrics and formulas recalculate fully. */
	structural?: boolean;
	sheet?: number;
	ranges?: CellRange[];
	/** Cells whose content changed, for incremental recalculation. */
	cells?: { sheet: number; row: number; col: number }[];
}

export interface ExternalHost {
	calculator: SessionCalculator;
	emit: (change: WorkbookChange) => void;
	/** True while a local edit step is open. */
	busy: () => boolean;
}

/** Runs `apply` (which mutates the workbook) outside the history. Returns whether it changed. */
export function applyExternal(
	host: ExternalHost,
	apply: () => ExternalChange | undefined,
): boolean {
	if (host.busy()) throw new Error('Cannot apply an external change inside an edit');
	const change = apply();
	if (!change) return false;
	const structural = change.structural ?? false;
	host.calculator.afterExternal(structural, change.cells ?? []);
	const out: WorkbookChange = {
		kind: change.kind ?? 'remote',
		label: change.label,
		structural,
		external: true,
	};
	if (change.sheet !== undefined) out.sheet = change.sheet;
	if (change.ranges?.length) out.ranges = change.ranges;
	host.emit(out);
	return true;
}
