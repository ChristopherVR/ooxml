import { TextSelection, type Command, type EditorState } from 'prosemirror-state';
import type { ResolvedPos } from 'prosemirror-model';
import { canExecuteTableCommand, executeTableCommand } from './table-commands';

/** Depth of the table cell holding the position, or -1 outside a table. */
function cellDepth($pos: ResolvedPos): number {
	for (let depth = $pos.depth; depth > 0; depth--)
		if ($pos.node(depth).type.name === 'tableCell') return depth;
	return -1;
}

/** Start position (before the cell node) of every cell of the table, in reading order. */
function cellPositions($pos: ResolvedPos, tableDepth: number): number[] {
	const positions: number[] = [];
	const table = $pos.node(tableDepth);
	let rowPos = $pos.start(tableDepth);
	table.forEach((row) => {
		let cellPos = rowPos + 1;
		row.forEach((cell) => {
			positions.push(cellPos);
			cellPos += cell.nodeSize;
		});
		rowPos += row.nodeSize;
	});
	return positions;
}

/** Select the whole text of the cell that starts at `pos`, as Word does when Tab enters a cell. */
function selectCell(state: EditorState, pos: number): TextSelection {
	const cell = state.doc.nodeAt(pos)!;
	const from = TextSelection.near(state.doc.resolve(pos + 1), 1);
	const to = TextSelection.near(state.doc.resolve(pos + cell.nodeSize - 1), -1);
	return TextSelection.create(state.doc, from.from, to.to);
}

/**
 * Word's Tab and Shift+Tab inside a table: move to the next or previous cell in reading order and
 * select its text. Tab in the last cell adds a row when the table's structure may change; a
 * structure-locked table (merged cells) keeps the caret where it is.
 */
export function cellTab(direction: 1 | -1): Command {
	return (state, dispatch, view) => {
		if (view && !view.editable) return false;
		const { $from } = state.selection;
		const depth = cellDepth($from);
		if (depth < 0) return false;
		const positions = cellPositions($from, depth - 2);
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
			const all = cellPositions(rebuilt.selection.$from, rebuiltDepth - 2);
			const columns = rebuilt.selection.$from.node(rebuiltDepth - 1).childCount;
			const first = all[all.length - columns];
			if (first !== undefined)
				view.dispatch(rebuilt.tr.setSelection(selectCell(rebuilt, first)).scrollIntoView());
			return true;
		}
		return true;
	};
}
