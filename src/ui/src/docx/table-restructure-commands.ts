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
import { closeHistory } from 'prosemirror-history';
import { Selection } from 'prosemirror-state';
import type { Node as ProseMirrorNode } from 'prosemirror-model';
import type { EditorView } from 'prosemirror-view';
import { schema } from './schema';
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
}

/** The innermost table around the selection, when the structure commands can edit it. */
export function restructureContext(view: EditorView): Context | undefined {
	const { $from, from, to } = view.state.selection;
	let depth = $from.depth;
	while (depth > 0 && $from.node(depth).type !== schema.nodes.table) depth--;
	if (depth <= 0 || $from.depth < depth + 2) return undefined;
	const table = $from.node(depth);
	if (!table.attrs.restructurable) return undefined;
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
			if (Math.max(from, pos + 1) <= Math.min(to, pos + slot.cell.nodeSize - 1))
				selected.push({ slot, index });
		}),
	);
	if (!selected.length) return undefined;
	const rect: GridRect = { firstRow: Infinity, lastRow: -1, firstCol: Infinity, lastCol: -1 };
	for (const { slot } of selected) {
		rect.firstRow = Math.min(rect.firstRow, slot.row);
		rect.lastRow = Math.max(rect.lastRow, slot.row + (Number(slot.cell.attrs.rowspan) || 1) - 1);
		rect.firstCol = Math.min(rect.firstCol, slot.col);
		rect.lastCol = Math.max(rect.lastCol, slot.col + slot.span - 1);
	}
	return {
		table,
		tablePos,
		nested: $from.node(depth - 1).type === schema.nodes.tableCell,
		skeleton,
		selected,
		rect,
	};
}

/** Whether the merged-table commands (not the rectangular-table ones) own this table. */
export function usesRestructure(ctx: Context | undefined, action: string): boolean {
	if (!ctx) return false;
	return (
		action === 'mergeCells' ||
		action === 'splitCell' ||
		ctx.nested ||
		ctx.table.attrs.structureEditable === false
	);
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
			const closed = closeRect(skeleton, rect);
			return closed && mergeCells(skeleton, closed, newId);
		}
		case 'splitCell': {
			const target = ctx.selected[0];
			if (!target || ctx.selected.length !== 1) return undefined;
			return splitCell(
				skeleton,
				target.slot.row,
				target.index,
				split?.rows ?? 1,
				split?.columns ?? 2,
				newId,
			);
		}
	}
}

export interface Split {
	rows: number;
	columns: number;
}

/** Whether `action` applies to the table around the selection (a dry run of the edit). */
export function canRestructure(ctx: Context, action: RestructureCommand): boolean {
	let serial = 0;
	const outcome = compute(ctx, action, () => `probe-${++serial}`);
	return outcome !== undefined;
}

/** The size of the cell the Split Cells dialog starts from, or undefined when none can be split. */
export function splitTarget(view: EditorView): { rows: number; span: number } | undefined {
	const ctx = restructureContext(view);
	const only = ctx?.selected.length === 1 ? ctx.selected[0] : undefined;
	if (!ctx || !only) return undefined;
	const grid = tableGrid(ctx.skeleton);
	const entry = grid?.rows[only.slot.row]?.[only.index];
	return grid && entry ? { rows: rowSpanOf(grid, entry), span: entry.span } : undefined;
}

/** Applies `action` to the table around the selection; false when it cannot be applied. */
export function executeRestructure(
	view: EditorView,
	action: RestructureCommand,
	newId: IdFactory,
	split?: Split,
): boolean {
	const ctx = restructureContext(view);
	if (!ctx || !view.editable) return false;
	const outcome = compute(ctx, action, newId, split);
	if (!outcome) return false;
	const { tablePos, table } = ctx;
	let replacement: ProseMirrorNode | undefined;
	if (outcome === 'table') replacement = schema.nodes.paragraph.create({ id: newId() });
	else replacement = nodeFromResult(table, outcome);
	if (!replacement) return false;
	const transaction = closeHistory(
		view.state.tr.replaceWith(tablePos, tablePos + table.nodeSize, replacement),
	);
	transaction.setMeta(RESTRUCTURE_TABLE_META, true);
	transaction.setSelection(
		Selection.near(transaction.doc.resolve(tablePos + (outcome === 'table' ? 1 : 2))),
	);
	view.dispatch(transaction.scrollIntoView());
	return true;
}
