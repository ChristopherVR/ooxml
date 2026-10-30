import { actionRun } from './action-formula';
import { controlFormulaDraggable } from './control-formula';
import type { Cells, Report, Row, Sheet } from './sheet';
import { number, sectionRows } from './sheet';
import type {
	VisioAction,
	VisioControl,
	VisioScratchRow,
	VisioShapeSections,
	VisioTabStop,
	VisioUserCell,
} from './shape-sections-types';

/** A shape of the page holds at most this many rows of each section; the rest are not modelled. */
const MAX_ROWS = 256;
const SCRATCH_CELLS = ['X', 'Y', 'A', 'B', 'C', 'D', 'E', 'F'] as const;

const ownRows = (local: Sheet, name: string) =>
	new Set(sectionRows(local, name).map((row) => row.index));
const text = (cells: Cells, name: string): string => cells.get(name)?.value ?? '';
const flag = (cells: Cells, name: string, report: Report): boolean =>
	number(cells, name, 0, report) !== 0;

function controlOf(row: Row, inherited: boolean, report: Report): VisioControl | undefined {
	const x = number(row.cells, 'X', NaN),
		y = number(row.cells, 'Y', NaN);
	if (!Number.isFinite(x) || !Number.isFinite(y)) return undefined;
	const xFree = number(row.cells, 'XCon', 0, report) === 0;
	const yFree = number(row.cells, 'YCon', 0, report) === 0;
	const drag = {
		x: xFree && controlFormulaDraggable(row.cells.get('X')?.formula, 'x'),
		y: yFree && controlFormulaDraggable(row.cells.get('Y')?.formula, 'y'),
	};
	const prompt = text(row.cells, 'Prompt');
	return {
		row: row.index,
		...(row.name === undefined ? {} : { name: row.name }),
		x,
		y,
		...(prompt ? { prompt } : {}),
		canGlue: flag(row.cells, 'CanGlue', report),
		drag,
		...(drag.x || drag.y
			? {}
			: {
					locked:
						xFree || yFree
							? 'Its position is computed by a formula that dragging cannot rewrite.'
							: 'Its position is fixed, or follows another handle.',
				}),
		inherited,
	};
}

function actionOf(row: Row, report: Report): VisioAction {
	const action = row.cells.get('Action')?.formula;
	const sortKey = text(row.cells, 'SortKey');
	return {
		row: row.index,
		...(row.name === undefined ? {} : { name: row.name }),
		// `&` marks the accelerator in a Visio menu; `&&` is a literal ampersand.
		menu: text(row.cells, 'Menu').replace(/&(&?)/g, '$1'),
		...(action === undefined || action === 'No Formula' ? {} : { action }),
		checked: flag(row.cells, 'Checked', report),
		disabled: flag(row.cells, 'Disabled', report),
		invisible: flag(row.cells, 'Invisible', report),
		beginGroup: flag(row.cells, 'BeginGroup', report),
		...(sortKey ? { sortKey } : {}),
		run: actionRun(action),
	};
}

/** Controls, Actions, User-defined and Scratch cells of a normalised (master-merged) shape sheet. */
export function shapeSections(
	sheet: Sheet,
	local: Sheet,
	inheritedSheet: boolean,
	report: Report,
): VisioShapeSections {
	const result: VisioShapeSections = {};
	const ownControls = new Set([...ownRows(local, 'Control'), ...ownRows(local, 'Controls')]);
	const controls = [...sectionRows(sheet, 'Control'), ...sectionRows(sheet, 'Controls')]
		.slice(0, MAX_ROWS)
		.flatMap((row) => {
			const control = controlOf(row, inheritedSheet && !ownControls.has(row.index), report);
			return control ? [control] : [];
		});
	if (controls.length) result.controls = controls;
	const actions = sectionRows(sheet, 'Actions')
		.slice(0, MAX_ROWS)
		.map((row) => actionOf(row, report));
	if (actions.length) result.actions = actions;
	const userCells: VisioUserCell[] = sectionRows(sheet, 'User')
		.slice(0, MAX_ROWS)
		.flatMap((row) => {
			const value = row.cells.get('Value');
			const name = row.name ?? row.index;
			if (!value || !name) return [];
			const prompt = text(row.cells, 'Prompt');
			return [
				{
					name,
					value: value.value ?? '',
					...(value.formula && value.formula !== 'No Formula' && value.formula !== 'Inh'
						? { formula: value.formula }
						: {}),
					...(prompt ? { prompt } : {}),
				},
			];
		});
	if (userCells.length) result.userCells = userCells;
	const scratch: VisioScratchRow[] = sectionRows(sheet, 'Scratch')
		.slice(0, MAX_ROWS)
		.map((row) => ({
			row: row.index,
			values: Object.fromEntries(
				SCRATCH_CELLS.flatMap((name) => {
					const value = number(row.cells, name, NaN);
					return Number.isFinite(value) ? [[name, value] as const] : [];
				}),
			),
		}))
		.filter((row) => Object.keys(row.values).length);
	if (scratch.length) result.scratch = scratch;
	return result;
}

const ALIGNMENTS: VisioTabStop['align'][] = ['left', 'center', 'right', 'decimal'];
/**
 * The tab stops of one Tabs row. A row lists its stops as numbered cell pairs (`Position1`,
 * `Alignment1`, `Position2`, ...). No Visio-authored file with explicit stops was available, so
 * the cell names follow the documented section; a row that names none has none.
 */
export function tabStops(cells: Cells, report: Report): VisioTabStop[] {
	const stops: VisioTabStop[] = [];
	for (const name of cells.keys()) {
		const match = /^(?:Tab)?Pos(?:ition)?(\d{1,3})$/i.exec(name);
		if (!match) continue;
		const position = number(cells, name, NaN, report);
		if (!Number.isFinite(position) || position < 0) continue;
		const align = number(
			cells,
			[...cells.keys()].find((key) =>
				new RegExp(`^(?:Tab)?Align(?:ment)?${match[1]}$`, 'i').test(key),
			) ?? '',
			0,
		);
		stops.push({ position, align: ALIGNMENTS[align] ?? 'left' });
	}
	return stops.sort((a, b) => a.position - b.position).slice(0, 64);
}
