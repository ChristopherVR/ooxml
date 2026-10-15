// Data > What-If Analysis: Goal Seek (with its status dialog) and Data Table.
import {
	displayText,
	formatAddress,
	getCell,
	parseInputCell,
	type GoalSeekRequest,
	type GoalSeekResult,
} from 'ooxml-core/xlsx';
import { target } from 'ooxml-core/xlsx/ui';
import type { EditorContext } from 'ooxml-core/xlsx/ui';
import { field, text, textInput } from './fields';
import { showDialog } from './frame';

/** Parses a typed number the way the dialog fields accept it (blank or text is NaN). */
const toNumber = (input: string): number => {
	const trimmed = input.trim().replace(/^=/, '');
	return trimmed === '' ? Number.NaN : Number(trimmed);
};

interface GoalSeekOutcome {
	request: GoalSeekRequest;
	result: GoalSeekResult;
}

/** Goal Seek: three fields, then a status dialog whose OK keeps the result and Cancel undoes it. */
export async function openGoalSeek(ctx: EditorContext): Promise<boolean | undefined> {
	const t = target(ctx);
	if (!t) return undefined;
	const setCell = textInput(ctx, formatAddress(t.active));
	const goal = textInput(ctx);
	const changing = textInput(ctx);
	const outcome = await showDialog<GoalSeekOutcome>(ctx, {
		name: 'goal-seek',
		heading: 'Goal Seek',
		body: [
			field(ctx, 'Set cell:', setCell),
			field(ctx, 'To value:', goal),
			field(ctx, 'By changing cell:', changing),
		],
		opened: () => goal.focus(),
		submit: () => {
			const set = parseInputCell(t.ws, setCell.value);
			if (!set) {
				ctx.toast(ctx.t('Reference is not valid.'), 'warning');
				setCell.focus();
				return undefined;
			}
			const by = parseInputCell(t.ws, changing.value);
			if (!by) {
				ctx.toast(ctx.t('Reference is not valid.'), 'warning');
				changing.focus();
				return undefined;
			}
			const request: GoalSeekRequest = {
				sheet: t.sheet,
				target: set,
				goal: toNumber(goal.value),
				changing: by,
			};
			return { request, result: t.session.goalSeek(request) };
		},
	});
	if (!outcome) return undefined;
	const { request, result } = outcome;
	const cell = formatAddress(request.target);
	const current = displayText(t.workbook, getCell(t.ws, request.target.row, request.target.col));
	const keep = await showDialog<boolean>(ctx, {
		name: 'goal-seek-status',
		heading: 'Goal Seek Status',
		body: [
			text(
				ctx,
				result.found
					? 'Goal Seeking with Cell {cell} found a solution.'
					: 'Goal Seeking with Cell {cell} may not have found a solution.',
				'xve-note',
				{ cell },
			),
			text(ctx, 'Target value: {value}', 'xve-note', { value: String(request.goal) }),
			text(ctx, 'Current value: {value}', 'xve-note', { value: current }),
		],
		submit: () => true,
	});
	// Cancel restores the original value of the changing cell.
	if (!keep) t.session.undo();
	return Boolean(keep);
}

/** Data Table: the row and column input cells for the selected range. */
export function openDataTable(ctx: EditorContext): Promise<boolean | undefined> {
	const t = target(ctx);
	if (!t) return Promise.resolve(undefined);
	const rowInput = textInput(ctx);
	const colInput = textInput(ctx);
	return showDialog<boolean>(ctx, {
		name: 'data-table',
		heading: 'Data Table',
		body: [
			field(ctx, 'Row input cell:', rowInput),
			field(ctx, 'Column input cell:', colInput),
			text(ctx, 'Select the input values and formulas first, then name the input cells.'),
		],
		opened: () => rowInput.focus(),
		submit: () => {
			t.session.createDataTable(t.sheet, {
				range: t.range,
				...(rowInput.value.trim() ? { rowInput: rowInput.value } : {}),
				...(colInput.value.trim() ? { colInput: colInput.value } : {}),
			});
			return true;
		},
	});
}
