import { TextSelection, type Command, type EditorState } from 'prosemirror-state';
import type { Node as ProseMirrorNode, ResolvedPos } from 'prosemirror-model';
import { canExecuteTableCommand, executeTableCommand } from './table-commands';

/** Depth of the innermost table cell holding the position, or -1 outside a table. */
function cellDepth($pos: ResolvedPos): number {
	for (let depth = $pos.depth; depth > 0; depth--)
		if ($pos.node(depth).type.name === 'tableCell') return depth;
	return -1;
}

/** Depth of the outermost table around the position, or -1 outside a table. */
function outerTableDepth($pos: ResolvedPos): number {
	for (let depth = 1; depth <= $pos.depth; depth++)
		if ($pos.node(depth).type.name === 'table') return depth;
	return -1;
}

/**
 * Start position (before the cell node) of every cell of the table, in reading order. A table
 * nested in a cell contributes its cells right after the cell that holds it, as Word's Tab does.
 */
function cellPositions(table: ProseMirrorNode, start: number): number[] {
	const positions: number[] = [];
	let rowPos = start;
	table.forEach((row) => {
		let cellPos = rowPos + 1;
		row.forEach((cell) => {
			positions.push(cellPos);
			let childPos = cellPos + 1;
			cell.forEach((child) => {
				if (child.type.name === 'table') positions.push(...cellPositions(child, childPos + 1));
				childPos += child.nodeSize;
			});
			cellPos += cell.nodeSize;
		});
		rowPos += row.nodeSize;
	});
	return positions;
}

/** Positions of the outermost table's cells for the selection, or none outside a table. */
function positionsAround($pos: ResolvedPos): number[] {
	const depth = outerTableDepth($pos);
	return depth < 0 ? [] : cellPositions($pos.node(depth), $pos.start(depth));
}

/** Select the whole text of the cell that starts at `pos`, as Word does when Tab enters a cell. */
function selectCell(state: EditorState, pos: number): TextSelection {
	const cell = state.doc.nodeAt(pos)!;
	const from = TextSelection.near(state.doc.resolve(pos + 1), 1);
	// A cell holding a nested table selects only the text before it; the nested cells follow.
	let end = pos + cell.nodeSize - 1;
	let childPos = pos + 1;
	cell.forEach((child) => {
		if (child.type.name === 'table') end = Math.min(end, childPos);
		childPos += child.nodeSize;
	});
	const to = TextSelection.near(state.doc.resolve(end), -1);
	return TextSelection.create(state.doc, from.from, Math.max(from.from, to.to));
}

/**
 * Word's Tab and Shift+Tab inside a table: move to the next or previous cell in reading order and
 * select its text. Tab in the last cell adds a row when the table's structure may change; a
 * structure-locked table (merged cells) keeps the caret where it is. Cells of nested tables take
 * part in the reading order, and Tab in the last nested cell continues in the enclosing table.
 */
export function cellTab(direction: 1 | -1): Command {
	return (state, dispatch, view) => {
		if (view && !view.editable) return false;
		const { $from } = state.selection;
		const depth = cellDepth($from);
		if (depth < 0) return false;
		const positions = positionsAround($from);
		const index = positions.indexOf($from.before(depth));
		const target = positions[index + direction];
		if (target !== undefined) {
			if (dispatch) dispatch(state.tr.setSelection(selectCell(state, target)).scrollIntoView());
			return true;
		}
		if (direction === 1 && view && canExecuteTableCommand(view, 'rowAfter') && dispatch) {
			executeTableCommand(view, 'rowAfter');
			// The table is rebuilt with the selection in its first cell; go to the new last row.
			const rebuilt = view.state;
			const rebuiltDepth = cellDepth(rebuilt.selection.$from);
			const all = positionsAround(rebuilt.selection.$from);
			const columns = rebuilt.selection.$from.node(rebuiltDepth - 1).childCount;
			const first = all[all.length - columns];
			if (first !== undefined)
				view.dispatch(rebuilt.tr.setSelection(selectCell(rebuilt, first)).scrollIntoView());
			return true;
		}
		return true;
	};
}
