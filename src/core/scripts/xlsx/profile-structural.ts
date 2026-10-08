// Times the structural edits of `<xlsx-editor>` on a loaded workbook (`profile-load.ts
// --structural`): insert and delete a row, insert a column, rename a sheet, Calculate Now (F9)
// and a full calculation (Ctrl+Alt+F9). Each edit is timed end to end on a live session, and
// again split into its reference transformation (a session that does not recalculate), the calc
// engine following it in place (graph update plus the formulas it queued), and, for comparison,
// a graph rebuild and a full re-evaluation (what every structural edit used to cost).
import { createEditSession, type EditSession } from '../../xlsx/edit/index';
import type { CalcEngine } from '../../xlsx/formula/index';
import { loadWorkbook } from '../../xlsx/load/index';
import type { Workbook } from '../../xlsx/model';

type Stage = <T>(label: string, run: () => T | Promise<T>) => Promise<T>;

interface Operation {
	label: string;
	apply: (session: EditSession, sheet: number) => void;
	/** What the session asks the engine after the edit (see `SessionCalculator`). */
	follow: (calc: CalcEngine, sheet: number, oldName: string, newName: string) => void;
}

/** The sheet with the most formulas: the edits are made there. */
function busiestSheet(workbook: Workbook): number {
	let best = 0;
	let most = -1;
	workbook.sheets.forEach((sheet, index) => {
		let count = 0;
		for (const row of sheet.rows.values())
			for (const cell of row.values()) if (cell.formula) count++;
		if (count > most) [best, most] = [index, count];
	});
	return best;
}

const shift =
	(axis: 'row' | 'col', at: number, count: number): Operation['follow'] =>
	(calc, sheet) =>
		calc.shiftCells(sheet, { axis, at, count });
const OPERATIONS: Operation[] = [
	{ label: 'insert row 6', apply: (s, i) => s.insertRows(i, 5, 1), follow: shift('row', 5, 1) },
	{ label: 'delete row 6', apply: (s, i) => s.deleteRows(i, 5, 1), follow: shift('row', 5, -1) },
	{
		label: 'insert column C',
		apply: (s, i) => s.insertColumns(i, 2, 1),
		follow: shift('col', 2, 1),
	},
	{
		label: 'delete column C',
		apply: (s, i) => s.deleteColumns(i, 2, 1),
		follow: shift('col', 2, -1),
	},
	{
		label: 'rename sheet',
		apply: (s, i) => s.renameSheet(i, `${s.workbook.sheets[i]?.name ?? 'S'} 2`),
		follow: (calc, _sheet, from, to) => calc.renameSheet(from, to),
	},
];

interface EngineInternals {
	invalidate(): void;
	build(): void;
}

const nameOf = (session: EditSession, sheet: number): string =>
	session.workbook.sheets[sheet]?.name ?? '';

/** Opens the workbook, calculates it once and returns a session ready to edit. */
async function ready(bytes: Uint8Array, path: string, recalc: boolean): Promise<EditSession> {
	const workbook = await loadWorkbook(bytes, { fileName: path });
	const session = createEditSession(workbook, { recalc, autoRowHeight: false });
	session.calc.recalculateAll();
	return session;
}

export async function profileStructural(
	bytes: Uint8Array,
	path: string,
	stage: Stage,
): Promise<void> {
	// The first edit after opening is structural, before idle preparation finished the graph.
	const opened = await loadWorkbook(bytes, { fileName: path });
	const first = createEditSession(opened, { autoRowHeight: false });
	let slices = 0;
	first.prepareCalculation({ timeRemaining: () => (slices++ < 200 ? 1 : 0) });
	await stage('insert row 6, graph half prepared (total)', () =>
		first.insertRows(busiestSheet(opened), 5, 1),
	);

	// End to end, as the editor runs them (one live session, edits in sequence).
	const live = await ready(bytes, path, true);
	const sheet = busiestSheet(live.workbook);
	for (const op of OPERATIONS) await stage(`${op.label} (total)`, () => op.apply(live, sheet));
	await stage('calculate now, F9 (total)', () => live.calculateNow());
	await stage('calculate full, Ctrl+Alt+F9 (total)', () => live.calculateNow({ full: true }));

	// The same edits split: transformation only, then a rebuild and a full recalculation.
	const split = await ready(bytes, path, false);
	const engine = split.calc as unknown as EngineInternals;
	for (const op of OPERATIONS) {
		const before = nameOf(split, sheet);
		await stage(`${op.label}: transform references`, () => op.apply(split, sheet));
		await stage(`${op.label}: follow in place + evaluate`, () => {
			op.follow(split.calc, sheet, before, nameOf(split, sheet));
			split.calc.recalculateFrom([]);
		});
		await stage(`${op.label}: rebuild graph`, () => {
			engine.invalidate();
			engine.build();
		});
		await stage(`${op.label}: rebuild + evaluate all`, () => {
			split.calc.invalidate();
			split.calc.recalculateAll();
		});
	}
}
