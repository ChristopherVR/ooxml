// Resolves a form control's linked cell and list range against the workbook, and gathers what its
// painter needs (state from the linked value, list items as displayed text).
import { getCell, getValue } from '../cells';
import { displayText } from '../display-text';
import type { FormControlObject, Workbook } from '../model';
import {
	controlState,
	radioGroup,
	type ControlState,
	type LinkedValue,
} from './form-control-state';
import { plainReference } from './ref-evaluator';

export interface ControlCell {
	sheet: number;
	row: number;
	col: number;
}

/** The single cell a `fmlaLink` such as `$C$1` or `Sheet2!$C$1` names; sheet defaults to `sheetIndex`. */
export function controlLinkCell(
	workbook: Workbook,
	sheetIndex: number,
	link: string | undefined,
): ControlCell | undefined {
	if (!link) return undefined;
	const found = plainReference(workbook, link.replace(/^=/, ''), sheetIndex);
	if (!found) return undefined;
	const { start } = found.range;
	return { sheet: found.sheet, row: start.row, col: start.col };
}

/** What a control's linked cell holds right now (`undefined` when unlinked or empty). */
export function controlLinkedValue(
	workbook: Workbook,
	sheetIndex: number,
	link: string | undefined,
): LinkedValue {
	const cell = controlLinkCell(workbook, sheetIndex, link);
	const sheet = cell ? workbook.sheets[cell.sheet] : undefined;
	if (!cell || !sheet) return undefined;
	const value = getValue(sheet, cell.row, cell.col);
	return value === null || typeof value === 'object' ? undefined : value;
}

/** The items of a drop-down or list box: the displayed text of each cell of its list range. */
export function controlListItems(
	workbook: Workbook,
	sheetIndex: number,
	control: FormControlObject,
): string[] {
	if (!control.listRange) return [];
	const found = plainReference(workbook, control.listRange, sheetIndex);
	const sheet = found ? workbook.sheets[found.sheet] : undefined;
	if (!found || !sheet) return [];
	const items: string[] = [];
	const { start, end } = found.range;
	for (let row = start.row; row <= Math.min(end.row, start.row + 999); row++)
		for (let col = start.col; col <= Math.min(end.col, start.col + 99); col++)
			items.push(displayText(workbook, getCell(sheet, row, col)));
	return items;
}

/** The state of a control in the sheet's drawing order, radio groups resolved. */
export function controlStateIn(
	workbook: Workbook,
	sheetIndex: number,
	controls: readonly FormControlObject[],
	control: FormControlObject,
): { state: ControlState; radioIndex: number; link: string | undefined } {
	if (control.controlType === 'radio') {
		const group = radioGroup(controls, control);
		const linked = controlLinkedValue(workbook, sheetIndex, group.linkedCell);
		return {
			state: controlState(
				group.linkedCell ? { ...control, linkedCell: group.linkedCell } : control,
				linked,
				group.index,
			),
			radioIndex: group.index,
			link: group.linkedCell,
		};
	}
	const linked = controlLinkedValue(workbook, sheetIndex, control.linkedCell);
	return { state: controlState(control, linked), radioIndex: 0, link: control.linkedCell };
}
