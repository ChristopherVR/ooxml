// Data > Reapply for sorting: re-runs the sort saved with a table, auto-filter or sheet.
import { type CellAddress, rangeContains } from '../address';
import { SOLID } from '../filter-criteria';
import { resolveColor } from '../layout/colors';
import type { SortCondition, SortState, Workbook, Worksheet } from '../model';
import { type EditContext, sheetAt } from './context';
import { resolveSortOptions, sortRows } from './sort';
import type { SortKey } from './sort-keys';

/** The saved sort that applies at `at`: its table's, then the filter's, then the sheet's. */
export function savedSortAt(sheet: Worksheet, at?: CellAddress): SortState | undefined {
	const table = at && sheet.tables.find((t) => rangeContains(t.range, at));
	if (table) return table.sortState;
	return sheet.autoFilter?.sortState ?? sheet.sortState;
}

function colorOf(workbook: Workbook, condition: SortCondition): string | undefined {
	const style = condition.style;
	const color = condition.sortBy === 'fontColor' ? style?.font?.color : SOLID(style?.fill);
	return resolveColor(color, workbook.theme)?.toUpperCase();
}

/** The sort levels a saved `sortState` stands for. */
export function keysOfSortState(workbook: Workbook, state: SortState): SortKey[] {
	return state.conditions.map((c) => {
		const key: SortKey = { col: state.columnSort ? c.ref.start.row : c.ref.start.col };
		if (c.descending) key.descending = true;
		if (c.sortBy && c.sortBy !== 'value') key.sortOn = c.sortBy;
		if (c.customList) key.customList = c.customList.split(',');
		if (c.sortBy === 'cellColor' || c.sortBy === 'fontColor') {
			const color = colorOf(workbook, c);
			if (color) key.color = color;
		}
		if (c.sortBy === 'icon' && c.iconSet !== undefined)
			key.icon = { set: c.iconSet, ...(c.iconId !== undefined ? { index: c.iconId } : {}) };
		return key;
	});
}

/** Applies `state` to its own data range (the header is not part of it). */
export function reapplySavedSort(ctx: EditContext, s: number, state: SortState): void {
	const sheet = sheetAt(ctx.workbook, s);
	const keys = keysOfSortState(ctx.workbook, state);
	if (!keys.length) return;
	const options = resolveSortOptions(ctx, s, keys, {
		...(state.columnSort ? { leftToRight: true } : {}),
		...(state.caseSensitive ? { caseSensitive: true } : {}),
	});
	sortRows(sheet, state.range, keys, false, options);
}
