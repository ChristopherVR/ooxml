// Insert and Delete (Ctrl++ / Ctrl+-): shift cells, or insert / delete entire rows or columns.
import { target } from '../commands/util.js';
import type { EditorContext } from '../context.js';
import { radios } from './fields.js';
import { showDialog } from './frame.js';

export type ShiftChoice = 'right' | 'down' | 'left' | 'up' | 'row' | 'column';

export function openCellShift(
	ctx: EditorContext,
	mode: 'insert' | 'delete',
): Promise<ShiftChoice | undefined> {
	const t = target(ctx);
	if (!t) return Promise.resolve(undefined);
	const options: ReadonlyArray<readonly [ShiftChoice, string]> =
		mode === 'insert'
			? [
					['right', 'Shift cells right'],
					['down', 'Shift cells down'],
					['row', 'Entire row'],
					['column', 'Entire column'],
				]
			: [
					['left', 'Shift cells left'],
					['up', 'Shift cells up'],
					['row', 'Entire row'],
					['column', 'Entire column'],
				];
	const r = t.range;
	const wide = r.end.col - r.start.col > r.end.row - r.start.row;
	const initial: ShiftChoice = mode === 'insert' ? (wide ? 'down' : 'right') : wide ? 'up' : 'left';
	const choice = radios(ctx, mode === 'insert' ? 'Insert' : 'Delete', options, initial);
	return showDialog<ShiftChoice>(ctx, {
		name: mode === 'insert' ? 'insert-cells' : 'delete-cells',
		heading: mode === 'insert' ? 'Insert' : 'Delete',
		body: choice.element,
		opened: () => choice.inputs.find((i) => i.checked)?.focus(),
		submit: () => {
			const value = choice.get() as ShiftChoice;
			const { session, sheet } = t;
			const rows = r.end.row - r.start.row + 1;
			const cols = r.end.col - r.start.col + 1;
			if (value === 'row') {
				if (mode === 'insert') session.insertRows(sheet, r.start.row, rows);
				else session.deleteRows(sheet, r.start.row, rows);
			} else if (value === 'column') {
				if (mode === 'insert') session.insertColumns(sheet, r.start.col, cols);
				else session.deleteColumns(sheet, r.start.col, cols);
			} else if (value === 'right' || value === 'down') session.insertCellsShift(sheet, r, value);
			else if (value === 'left' || value === 'up') session.deleteCellsShift(sheet, r, value);
			return value;
		},
	});
}

/** Data > Group / Ungroup on a selection that is neither whole rows nor whole columns. */
export function openOutlineAxis(
	ctx: EditorContext,
	mode: 'group' | 'ungroup',
): Promise<'rows' | 'cols' | undefined> {
	const heading = mode === 'group' ? 'Group' : 'Ungroup';
	const choice = radios(
		ctx,
		heading,
		[
			['rows', 'Rows'],
			['cols', 'Columns'],
		],
		'rows',
	);
	return showDialog<'rows' | 'cols'>(ctx, {
		name: mode,
		heading,
		body: choice.element,
		opened: () => choice.inputs.find((i) => i.checked)?.focus(),
		submit: () => (choice.get() === 'cols' ? 'cols' : 'rows'),
	});
}
