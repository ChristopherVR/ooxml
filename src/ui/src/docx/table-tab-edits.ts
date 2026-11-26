import type { DocumentModel, Table, TableStyleCatalog } from 'ooxml-core/docx';
import { resolveTableCellMargins } from 'ooxml-core/docx';
import { paragraphAttrs, tableCellViews, tableViewAttrs } from 'ooxml-core/docx/ui';
import { closeHistory } from 'prosemirror-history';
import type { Node } from 'prosemirror-model';
import { Selection, SelectionRange, TextSelection, type Transaction } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { runToInlineNodes } from './run-adapter';
import { schema } from './schema';
import { fromRanges } from './table-cell-selection';
import { tableCells, tableTabContext } from './table-tab-context';
import type { TableTabAction } from './table-ribbon';

export type TableContext = NonNullable<ReturnType<typeof tableTabContext>>;

export function currentTable(model: DocumentModel, context: TableContext): Table | undefined {
	return model.blocks.find(
		(block): block is Table => block.type === 'table' && block.id === context.node.attrs.id,
	);
}

/** Applies only document attributes that changed. Keep the caret and all inline nodes intact. */
export function patchTableNode(
	tr: Transaction,
	pos: number,
	node: Node,
	table: Table,
	styles?: TableStyleCatalog,
	theme?: DocumentModel['theme'],
): Transaction {
	const assign = (at: number, previous: Node, attrs: Record<string, unknown>) => {
		for (const [key, value] of Object.entries(attrs))
			if (JSON.stringify(previous.attrs[key]) !== JSON.stringify(value))
				tr.setNodeAttribute(at, key, value);
	};
	assign(pos, node, {
		...tableViewAttrs(table),
		tablePropertiesKnown: true,
		tableStyleId: table.style ?? '',
		tableLook: table.look ?? null,
		tableLayout: table.layout ?? null,
		grid: table.grid ? JSON.stringify(table.grid) : null,
		inheritedCellMargins: JSON.stringify(
			resolveTableCellMargins(table.style ? { style: table.style } : {}, styles),
		),
	});
	const views = tableCellViews(table, styles, theme);
	node.forEach((row, rowOffset, rowIndex) => {
		const rowPos = pos + 1 + rowOffset;
		const properties = table.rowProperties?.[rowIndex] ?? null;
		assign(rowPos, row, {
			properties,
			heightPx: properties?.heightTwips ? properties.heightTwips / 15 : null,
		});
		row.forEach((cell, offset, column) => {
			const view = views[rowIndex]?.[column];
			if (view)
				assign(rowPos + 1 + offset, cell, {
					...view.attrs,
					widthPercent: view.cell.widthPercent ?? null,
				});
		});
	});
	return tr;
}

export function selectTableParts(
	view: EditorView,
	context: TableContext,
	kind: Extract<TableTabAction, { key: 'select' }>['value'],
): void {
	const { node, pos, top, firstColumn } = context;
	const { doc } = view.state;
	const cells = tableCells(node, pos).filter(
		({ row, column }) =>
			kind === 'table' ||
			(kind === 'row' && row === top) ||
			(kind === 'column' && column === firstColumn) ||
			(kind === 'cell' && row === top && column === firstColumn),
	);
	const ranges = cells.map(
		(cell) =>
			new SelectionRange(
				Selection.near(doc.resolve(cell.pos + 1), 1).$from,
				Selection.near(doc.resolve(cell.pos + cell.node.nodeSize - 1), -1).$from,
			),
	);
	const first = ranges[0];
	const last = ranges.at(-1);
	if (!first || !last) return;
	const selection =
		kind === 'column' ? fromRanges(ranges, doc) : TextSelection.between(first.$from, last.$to);
	view.dispatch(view.state.tr.setSelection(selection).scrollIntoView());
}

export function alignTableCells(
	view: EditorView,
	context: TableContext,
	action: Extract<TableTabAction, { key: 'align' }>,
): void {
	const tr = view.state.tr;
	for (const cell of tableCells(context.node, context.pos)) {
		if (
			!view.state.selection.ranges.some(
				({ $from, $to }) => $from.pos < cell.pos + cell.node.nodeSize && $to.pos > cell.pos,
			)
		)
			continue;
		tr.setNodeAttribute(cell.pos, 'verticalAlign', action.vertical);
		cell.node.forEach((child, offset) => {
			if (child.type.name !== 'paragraph') return;
			tr.setNodeAttribute(cell.pos + 1 + offset, 'align', action.horizontal);
			tr.setNodeAttribute(cell.pos + 1 + offset, 'justification', action.horizontal);
		});
	}
	if (tr.docChanged) view.dispatch(closeHistory(tr));
}

export function replaceTableWithText(
	view: EditorView,
	context: TableContext,
	paragraphs: Table['rows'][number][number]['paragraphs'],
) {
	const originals = new Map<string, Node>();
	context.node.descendants((node) => {
		if (node.type.name === 'paragraph') originals.set(String(node.attrs.id), node);
	});
	const nodes = paragraphs.map((paragraph) => {
		const original = originals.get(paragraph.id);
		const content = paragraph.runs.flatMap((run) => runToInlineNodes(run));
		return schema.nodes.paragraph.create(original?.attrs ?? paragraphAttrs(paragraph), content);
	});
	const tr = view.state.tr.replaceWith(context.pos, context.pos + context.node.nodeSize, nodes);
	tr.setSelection(Selection.near(tr.doc.resolve(context.pos), 1));
	view.dispatch(closeHistory(tr).scrollIntoView());
}

export function reorderTableRows(
	view: EditorView,
	context: TableContext,
	original: Table,
	sorted: Table,
): void {
	const nodes = new Map(original.rows.map((row, index) => [row, context.node.child(index)]));
	const rows = sorted.rows.flatMap((row) => {
		const node = nodes.get(row);
		return node ? [node] : [];
	});
	if (
		rows.length !== original.rows.length ||
		rows.every((row, index) => row === context.node.child(index))
	)
		return;
	const tr = view.state.tr.replaceWith(
		context.pos + 1,
		context.pos + context.node.nodeSize - 1,
		rows,
	);
	tr.setSelection(Selection.near(tr.doc.resolve(context.pos + 1), 1));
	view.dispatch(closeHistory(tr).scrollIntoView());
}
