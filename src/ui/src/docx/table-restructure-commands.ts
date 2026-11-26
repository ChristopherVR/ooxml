// Table commands for merged and nested tables, and the merge and split commands for any table.
// The table node is turned into the core model, the core structure edit (ooxml-core table-structure)
// runs on it, and the result replaces the table in one undoable step.
import {
	RESTRUCTURE_TABLE_META,
	closeRect,
	deleteColumns,
	deleteRow,
	insertColumn,
	insertRow,
	mergeCells,
	rowSpanOf,
	splitCell,
	tableGrid,
	type GridRect,
	type IdFactory,
} from 'ooxml-core/docx/ui';
import type { Table } from 'ooxml-core/docx';
import { Selection, TextSelection, type Transaction } from 'prosemirror-state';
import type { Node as ProseMirrorNode } from 'prosemirror-model';
import type { EditorView } from 'prosemirror-view';
import { schema } from './schema';
import { selectCellNodes } from './table-cell-selection';
import { dispatchTableEdit, tableCaret } from './table-edit-transaction';
import { nodeFromResult, skeletonOf, slotsOf, type Slot } from './table-restructure';

export type RestructureCommand =
	| 'rowBefore'
	| 'rowAfter'
	| 'deleteRow'
	| 'columnBefore'
	| 'columnAfter'
	| 'deleteColumn'
	| 'mergeCells'
	| 'splitCell';

interface Context {
	table: ProseMirrorNode;
	tablePos: number;
	nested: boolean;
	skeleton: Table;
	/** The cells the selection touches, with their row and position in the row. */
	selected: { slot: Slot; index: number }[];
	rect: GridRect;
	mergeRect: GridRect;
}

/** The innermost table around the selection, when the structure commands can edit it. */
export function restructureContext(view: EditorView, includePlain = false): Context | undefined {
	const { $from, ranges } = view.state.selection;
	let depth = $from.depth;
	while (depth > 0 && $from.node(depth).type !== schema.nodes.table) depth--;
	if (depth <= 0 || $from.depth < depth + 2) return undefined;
	const table = $from.node(depth);
	if (!table.attrs.restructurable && (!includePlain || table.attrs.structureEditable === false))
		return undefined;
	const last = ranges.at(-1)?.$to;
	if (!last || last.depth < depth + 2 || last.node(depth) !== table) return undefined;
	const tablePos = $from.before(depth);
	const skeleton = skeletonOf(table);
	if (!tableGrid(skeleton)) return undefined;
	const slots = slotsOf(table);
	const positions = new Map<ProseMirrorNode, number>();
	table.forEach((row, rowOffset) =>
		row.forEach((cell, cellOffset) =>
			positions.set(cell, tablePos + 1 + rowOffset + 1 + cellOffset),
		),
	);
	const selected: Context['selected'] = [];
	slots.forEach((row) =>
		row.forEach((slot, index) => {
			const pos = positions.get(slot.cell);
			if (slot.offset || pos === undefined) return;
			if (ranges.some(({ $from, $to }) => $from.pos < pos + slot.cell.nodeSize && $to.pos > pos))
				selected.push({ slot, index });
		}),
	);
	const first = selected[0];
	const final = selected.at(-1);
	if (!first || !final) return undefined;
	const rect = boundsOf(selected);
	return {
		table,
		tablePos,
		nested: $from.node(depth - 1).type === schema.nodes.tableCell,
		skeleton,
		selected,
		rect,
		// A drag or Shift+Arrow range defines the corners of the chosen rectangle.
		mergeRect: view.state.selection instanceof TextSelection ? boundsOf([first, final]) : rect,
	};
}

function boundsOf(selected: Context['selected']): GridRect {
	const rect: GridRect = { firstRow: Infinity, lastRow: -1, firstCol: Infinity, lastCol: -1 };
	for (const { slot } of selected) {
		rect.firstRow = Math.min(rect.firstRow, slot.row);
		rect.lastRow = Math.max(rect.lastRow, slot.row + (Number(slot.cell.attrs.rowspan) || 1) - 1);
		rect.firstCol = Math.min(rect.firstCol, slot.col);
		rect.lastCol = Math.max(rect.lastCol, slot.col + slot.span - 1);
	}
	return rect;
}

/** Whether the merged-table commands (not the rectangular-table ones) own this table. */
export function usesRestructure(ctx: Context | undefined): boolean {
	return Boolean(ctx && (ctx.nested || ctx.table.attrs.structureEditable === false));
}

type Outcome = Table | 'table' | undefined;

function compute(
	ctx: Context,
	action: RestructureCommand,
	newId: IdFactory,
	split?: Split,
): Outcome {
	const { skeleton, rect } = ctx;
	switch (action) {
		case 'rowBefore':
			return insertRow(skeleton, rect.firstRow, 'above', newId);
		case 'rowAfter':
			return insertRow(skeleton, rect.lastRow, 'below', newId);
		case 'deleteRow': {
			let table: Table = skeleton;
			for (let r = rect.lastRow; r >= rect.firstRow; r--) {
				const next = deleteRow(table, r);
				if (!next) return 'table';
				table = next;
			}
			return table;
		}
		// Word measures a single selected cell from its first column, a merged one included.
		case 'columnBefore':
			return insertColumn(skeleton, rect.firstCol, 'left', newId);
		case 'columnAfter':
			return insertColumn(
				skeleton,
				ctx.selected.length === 1 ? rect.firstCol : rect.lastCol,
				'right',
				newId,
			);
		case 'deleteColumn':
			// Taking out every column takes the table with it; a refused edit stays refused.
			return rect.firstCol === 0 && rect.lastCol >= (skeleton.grid?.length ?? Infinity) - 1
				? 'table'
				: deleteColumns(skeleton, rect.firstCol, rect.lastCol);
		case 'mergeCells': {
			if (ctx.selected.length < 2) return undefined;
			const closed = closeRect(skeleton, ctx.mergeRect);
			return closed && mergeCells(skeleton, closed, newId);
		}
		case 'splitCell': {
			const target = ctx.selected[0];
			if (!target || ctx.selected.length !== 1) return undefined;
			const size = split ?? splitDimensions(ctx);
			return (
				size && splitCell(skeleton, target.slot.row, target.index, size.rows, size.columns, newId)
			);
		}
	}
}

export interface Split {
	rows: number;
	columns: number;
}

function selectionAfterRestructure(
	tr: Transaction,
	ctx: Context,
	action: RestructureCommand,
	table: ProseMirrorNode,
	outcome: Table,
	caret: EditorView['state']['selection']['$head'],
	split?: Split,
): Selection {
	const { tablePos, rect, selected } = ctx;
	if (
		action === 'rowBefore' ||
		action === 'rowAfter' ||
		action === 'columnBefore' ||
		action === 'columnAfter'
	)
		return tableCaret(tr, table, tablePos, caret);
	const slots = slotsOf(table);
	const cellAt = (row: number, column: number) =>
		slots[row]?.find((slot) => slot.col <= column && column < slot.col + slot.span)?.cell;
	if (action === 'deleteRow' && rect.firstRow >= table.childCount)
		return Selection.near(tr.doc.resolve(tablePos + table.nodeSize), 1);
	if (action === 'deleteRow' || action === 'deleteColumn') {
		const cell =
			action === 'deleteRow'
				? cellAt(rect.firstRow, 0)
				: cellAt(0, Math.min(rect.firstCol, (outcome.grid?.length ?? 1) - 1));
		let pos = tablePos + 2;
		table.forEach((row, rowOffset) =>
			row.forEach((node, offset) => {
				if (node === cell) pos = tablePos + rowOffset + offset + 2;
			}),
		);
		return Selection.near(tr.doc.resolve(pos + 1), 1);
	}
	const cells = new Set<ProseMirrorNode>();
	const add = (cell: ProseMirrorNode | undefined) => {
		if (cell) cells.add(cell);
	};
	if (action === 'mergeCells') {
		const closed = closeRect(ctx.skeleton, ctx.mergeRect);
		if (closed) add(cellAt(closed.firstRow, closed.firstCol));
	} else {
		const target = selected[0];
		const first = target && slots[rect.firstRow]?.[target.index];
		if (first) {
			const last = slots[first.row]?.[target.index + (split?.columns ?? 2) - 1];
			for (let r = first.row; r < first.row + (split?.rows ?? 1); r++)
				for (const slot of slots[r] ?? [])
					if (slot.col >= first.col && slot.col <= (last?.col ?? first.col)) add(slot.cell);
		}
	}
	const selection = selectCellNodes(tr.doc, table, tablePos, cells);
	return action === 'mergeCells' ? Selection.near(selection.$from) : selection;
}

/** Whether `action` applies to the table around the selection (a dry run of the edit). */
export function canRestructure(ctx: Context, action: RestructureCommand, split?: Split): boolean {
	let serial = 0;
	const outcome = compute(ctx, action, () => `probe-${++serial}`, split);
	return outcome !== undefined;
}

/** The size of the cell the Split Cells dialog starts from, or undefined when none can be split. */
export function splitTarget(view: EditorView): Split | undefined {
	const ctx = restructureContext(view, true);
	return ctx && splitDimensions(ctx);
}

function splitDimensions(ctx: Context): Split | undefined {
	const only = ctx.selected.length === 1 ? ctx.selected[0] : undefined;
	if (!only) return;
	const grid = tableGrid(ctx.skeleton);
	const entry = grid?.rows[only.slot.row]?.[only.index];
	if (!grid || !entry) return;
	const rows = rowSpanOf(grid, entry);
	return { rows, columns: entry.span > 1 || rows > 1 ? entry.span : 2 };
}

/** Applies `action` to the table around the selection; false when it cannot be applied. */
export function executeRestructure(
	view: EditorView,
	action: RestructureCommand,
	newId: IdFactory,
	split?: Split,
): boolean {
	const ctx = restructureContext(view, action === 'mergeCells' || action === 'splitCell');
	if (!ctx || !view.editable) return false;
	const outcome = compute(ctx, action, newId, split);
	if (!outcome) return false;
	const { tablePos, table } = ctx;
	let replacement: ProseMirrorNode | undefined;
	if (outcome === 'table') replacement = schema.nodes.paragraph.create({ id: newId() });
	else replacement = nodeFromResult(table, outcome);
	if (!replacement) return false;
	const transaction = view.state.tr.replaceWith(tablePos, tablePos + table.nodeSize, replacement);
	transaction.setMeta(RESTRUCTURE_TABLE_META, true);
	transaction.setSelection(
		outcome === 'table'
			? Selection.near(transaction.doc.resolve(tablePos + 1))
			: selectionAfterRestructure(
					transaction,
					ctx,
					action,
					replacement,
					outcome,
					view.state.selection.$head,
					split ?? (action === 'splitCell' ? splitDimensions(ctx) : undefined),
				),
	);
	dispatchTableEdit(view, transaction.scrollIntoView());
	return true;
}
