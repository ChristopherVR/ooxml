import { expectDefined } from 'ooxml-core/docx/ui';
import { Fragment, type Node as ProseMirrorNode } from 'prosemirror-model';
import { closeHistory } from 'prosemirror-history';
import { Selection } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { schema } from './schema';
import { selectCellNodes } from './table-cell-selection';
import {
	canRestructure,
	executeRestructure,
	restructureContext,
	usesRestructure,
	type Split,
} from './table-restructure-commands';

export type TableCommand =
	| 'rowBefore'
	| 'rowAfter'
	| 'deleteRow'
	| 'columnBefore'
	| 'columnAfter'
	| 'deleteColumn'
	| 'mergeCells'
	| 'splitCell'
	| 'deleteTable';

type Context = {
	table: ProseMirrorNode;
	tablePos: number;
	tableDepth: number;
	rowIndex: number;
	columnIndex: number;
};

function context(view: EditorView): Context | null {
	const { $from } = view.state.selection;
	let tableDepth = -1;
	for (let depth = $from.depth; depth > 0; depth--) {
		if ($from.node(depth).type === schema.nodes.table) {
			tableDepth = depth;
			break;
		}
	}
	if (tableDepth < 0 || $from.depth < tableDepth + 2) return null;
	const table = $from.node(tableDepth);
	const width = table.childCount ? table.child(0).childCount : 0;
	if (!width || table.childCount < 1) return null;
	for (let row = 0; row < table.childCount; row++)
		if (table.child(row).childCount !== width) return null;
	return {
		table,
		tablePos: $from.before(tableDepth),
		tableDepth,
		rowIndex: $from.index(tableDepth),
		columnIndex: $from.index(tableDepth + 1),
	};
}

export function paragraphIdFactory(view: EditorView): () => string {
	const ids = new Set<string>();
	view.state.doc.descendants((node) => {
		if (typeof node.attrs.id === 'string' && node.attrs.id) ids.add(node.attrs.id);
	});
	let serial = 0;
	return () => {
		let id = '';
		do id = `dve-cell-${++serial}`;
		while (ids.has(id));
		ids.add(id);
		return id;
	};
}

function emptyCell(nextId: () => string): ProseMirrorNode {
	return schema.nodes.tableCell.create(null, schema.nodes.paragraph.create({ id: nextId() }));
}

function replacedTable(
	view: EditorView,
	ctx: Context,
	rows: ProseMirrorNode[],
	target: { row: number } | { column: number } | { caret: [number, number] },
): void {
	const next = ctx.table.copy(Fragment.fromArray(rows));
	const transaction = closeHistory(
		view.state.tr.replaceWith(ctx.tablePos, ctx.tablePos + ctx.table.nodeSize, next),
	);
	if ('caret' in target) {
		const [row, column] = target.caret;
		const pos =
			row >= rows.length
				? ctx.tablePos + next.nodeSize
				: ctx.tablePos + 2 + cellOffset(next, row, column);
		transaction.setSelection(Selection.near(transaction.doc.resolve(pos), 1));
	} else {
		const cells = new Set<ProseMirrorNode>();
		if ('row' in target) next.child(target.row).forEach((cell) => cells.add(cell));
		else next.forEach((row) => cells.add(row.child(target.column)));
		transaction.setSelection(selectCellNodes(transaction.doc, next, ctx.tablePos, cells));
	}
	view.dispatch(transaction.scrollIntoView());
}

function cellOffset(table: ProseMirrorNode, row: number, column: number): number {
	let offset = 0;
	for (let r = 0; r < row; r++) offset += table.child(r).nodeSize;
	for (let c = 0; c < column; c++) offset += table.child(row).child(c).nodeSize;
	return offset;
}

function selectedBounds(view: EditorView, ctx: Context) {
	const { table, tablePos, rowIndex, columnIndex } = ctx;
	const bounds = {
		firstRow: rowIndex,
		lastRow: rowIndex,
		firstCol: columnIndex,
		lastCol: columnIndex,
	};
	table.forEach((row, rowOffset, r) =>
		row.forEach((cell, cellOffset, c) => {
			const pos = tablePos + rowOffset + cellOffset + 2;
			if (
				!view.state.selection.ranges.some(
					({ $from, $to }) => $from.pos < pos + cell.nodeSize && $to.pos > pos,
				)
			)
				return;
			bounds.firstRow = Math.min(bounds.firstRow, r);
			bounds.lastRow = Math.max(bounds.lastRow, r);
			bounds.firstCol = Math.min(bounds.firstCol, c);
			bounds.lastCol = Math.max(bounds.lastCol, c);
		}),
	);
	return bounds;
}

function replaceWithParagraph(view: EditorView, ctx: Context): void {
	const nextId = paragraphIdFactory(view);
	const paragraph = schema.nodes.paragraph.create({ id: nextId() });
	const transaction = closeHistory(
		view.state.tr.replaceWith(ctx.tablePos, ctx.tablePos + ctx.table.nodeSize, paragraph),
	);
	transaction.setSelection(Selection.near(transaction.doc.resolve(ctx.tablePos + 1)));
	view.dispatch(transaction.scrollIntoView());
}

/** Reports whether the current selection can receive the requested table edit. */
export function canExecuteTableCommand(
	view: EditorView,
	action: TableCommand,
	split?: Split,
): boolean {
	if (action === 'deleteTable') return canDeleteTable(view);
	const needsGrid = action === 'mergeCells' || action === 'splitCell';
	const merged = restructureContext(view, needsGrid);
	if (needsGrid || usesRestructure(merged))
		return Boolean(view.editable && merged && canRestructure(merged, action, split));
	const ctx = context(view);
	// A nested table keeps the grid it was authored with; only its text is editable.
	const nested =
		ctx && view.state.selection.$from.node(ctx.tableDepth - 1).type.name === 'tableCell';
	return Boolean(view.editable && ctx && !nested && ctx.table.attrs.structureEditable !== false);
}

/** Any table that is not nested in a cell can be deleted whole, merged or not. */
function canDeleteTable(view: EditorView): boolean {
	const ctx = context(view) ?? tableAround(view);
	const nested =
		ctx && view.state.selection.$from.node(ctx.tableDepth - 1).type === schema.nodes.tableCell;
	return Boolean(view.editable && ctx && !nested);
}

/** The table around the selection when it is not rectangular, for deleting it whole. */
function tableAround(view: EditorView): Context | null {
	const { $from } = view.state.selection;
	for (let depth = $from.depth; depth > 0; depth--)
		if ($from.node(depth).type === schema.nodes.table && depth > 0 && $from.depth >= depth + 2)
			return {
				table: $from.node(depth),
				tablePos: $from.before(depth),
				tableDepth: depth,
				rowIndex: $from.index(depth),
				columnIndex: $from.index(depth + 1),
			};
	return null;
}

/**
 * Applies a structural command to the table containing the selection: rows and columns of a
 * rectangular table directly, merged and nested tables, and merge and split, through the
 * merge-aware editor. `split` sizes Split Cells (it defaults to two columns).
 */
export function executeTableCommand(
	view: EditorView,
	action: TableCommand,
	split?: Split,
): boolean {
	if (!canExecuteTableCommand(view, action, split)) return false;
	if (
		action === 'mergeCells' ||
		action === 'splitCell' ||
		(action !== 'deleteTable' && usesRestructure(restructureContext(view)))
	)
		return executeRestructure(view, action, paragraphIdFactory(view), split);
	const ctx = context(view) ?? tableAround(view);
	if (!ctx) return false;
	if (action === 'deleteTable') {
		replaceWithParagraph(view, ctx);
		return true;
	}

	const rows = Array.from({ length: ctx.table.childCount }, (_, index) => ctx.table.child(index));
	const { firstRow, lastRow, firstCol, lastCol } = selectedBounds(view, ctx);
	const width = expectDefined(rows[0], 'first table row').childCount;
	const nextId = paragraphIdFactory(view);
	if (action === 'deleteRow') {
		if (lastRow - firstRow + 1 === rows.length) replaceWithParagraph(view, ctx);
		else
			replacedTable(
				view,
				ctx,
				rows.filter((_, index) => index < firstRow || index > lastRow),
				{ caret: [firstRow, 0] },
			);
		return true;
	}
	if (action === 'rowBefore' || action === 'rowAfter') {
		const at = action === 'rowAfter' ? lastRow + 1 : firstRow;
		const added = schema.nodes.tableRow.create(
			null,
			Array.from({ length: width }, () => emptyCell(nextId)),
		);
		rows.splice(at, 0, added);
		replacedTable(view, ctx, rows, { row: at });
		return true;
	}

	if (action === 'deleteColumn') {
		if (lastCol - firstCol + 1 === width) replaceWithParagraph(view, ctx);
		else
			replacedTable(
				view,
				ctx,
				rows.map((row) =>
					row.copy(
						Fragment.fromArray(
							Array.from({ length: row.childCount }, (_, index) => row.child(index)).filter(
								(_, index) => index < firstCol || index > lastCol,
							),
						),
					),
				),
				{ caret: [0, Math.min(firstCol, width - (lastCol - firstCol + 1) - 1)] },
			);
		return true;
	}

	if (action !== 'columnBefore' && action !== 'columnAfter') return false;
	const at = action === 'columnAfter' ? lastCol + 1 : firstCol;
	replacedTable(
		view,
		ctx,
		rows.map((row) => {
			const cells = Array.from({ length: row.childCount }, (_, index) => row.child(index));
			cells.splice(at, 0, emptyCell(nextId));
			return row.copy(Fragment.fromArray(cells));
		}),
		{ column: at },
	);
	return true;
}
