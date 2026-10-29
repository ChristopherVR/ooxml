import { expectDefined } from 'ooxml-core/docx/ui';
import { Fragment, type Node as ProseMirrorNode } from 'prosemirror-model';
import { closeHistory } from 'prosemirror-history';
import { Selection } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { schema } from './schema';
import {
	canRestructure,
	executeRestructure,
	restructureContext,
	usesRestructure,
	type RestructureCommand,
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

function replacedTable(view: EditorView, ctx: Context, rows: ProseMirrorNode[]): void {
	const next = ctx.table.copy(Fragment.fromArray(rows));
	const transaction = closeHistory(
		view.state.tr.replaceWith(ctx.tablePos, ctx.tablePos + ctx.table.nodeSize, next),
	);
	transaction.setSelection(Selection.near(transaction.doc.resolve(ctx.tablePos + 2)));
	view.dispatch(transaction.scrollIntoView());
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
export function canExecuteTableCommand(view: EditorView, action: TableCommand): boolean {
	if (action === 'deleteTable') return canDeleteTable(view);
	const merged = restructureContext(view);
	if (usesRestructure(merged, action))
		return Boolean(view.editable && merged && canRestructure(merged, action as RestructureCommand));
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
	if (!canExecuteTableCommand(view, action)) return false;
	const merged = restructureContext(view);
	if (action !== 'deleteTable' && usesRestructure(merged, action))
		return executeRestructure(view, action as RestructureCommand, paragraphIdFactory(view), split);
	const ctx = context(view) ?? tableAround(view);
	if (!ctx) return false;
	if (action === 'deleteTable') {
		replaceWithParagraph(view, ctx);
		return true;
	}

	const rows = Array.from({ length: ctx.table.childCount }, (_, index) => ctx.table.child(index));
	const nextId = paragraphIdFactory(view);
	if (action === 'deleteRow') {
		if (rows.length === 1) replaceWithParagraph(view, ctx);
		else
			replacedTable(
				view,
				ctx,
				rows.filter((_, index) => index !== ctx.rowIndex),
			);
		return true;
	}
	if (action === 'rowBefore' || action === 'rowAfter') {
		const at = ctx.rowIndex + (action === 'rowAfter' ? 1 : 0);
		const added = schema.nodes.tableRow.create(
			null,
			Array.from({ length: expectDefined(rows[0], 'first table row').childCount }, () =>
				emptyCell(nextId),
			),
		);
		rows.splice(at, 0, added);
		replacedTable(view, ctx, rows);
		return true;
	}

	if (action === 'deleteColumn') {
		if (expectDefined(rows[0], 'first table row').childCount === 1) replaceWithParagraph(view, ctx);
		else
			replacedTable(
				view,
				ctx,
				rows.map((row) =>
					row.copy(
						Fragment.fromArray(
							Array.from({ length: row.childCount }, (_, index) => row.child(index)).filter(
								(_, index) => index !== ctx.columnIndex,
							),
						),
					),
				),
			);
		return true;
	}

	const at = ctx.columnIndex + (action === 'columnAfter' ? 1 : 0);
	replacedTable(
		view,
		ctx,
		rows.map((row) => {
			const cells = Array.from({ length: row.childCount }, (_, index) => row.child(index));
			cells.splice(at, 0, emptyCell(nextId));
			return row.copy(Fragment.fromArray(cells));
		}),
	);
	return true;
}
