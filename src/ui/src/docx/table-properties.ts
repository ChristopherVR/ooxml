import { signedTwips, type Table, type TableCell, type TableCellMargins } from 'ooxml-core/docx';
import { dispatchTableEdit } from './table-edit-transaction';
import type { EditorState } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';

export type TableRowProperties = NonNullable<Table['rowProperties']>[number];
export type RowPatch = { [K in keyof TableRowProperties]?: TableRowProperties[K] | undefined };

const WORD_FALLBACK_MARGINS: TableCellMargins = {
	left: signedTwips(108),
	right: signedTwips(108),
	top: signedTwips(0),
	bottom: signedTwips(0),
};

/** The margins a table inherits from its style chain (layout's resolution), for the dialogs. */
function inheritedCellMargins(json: unknown): TableCellMargins {
	return typeof json === 'string' ? (JSON.parse(json) as TableCellMargins) : WORD_FALLBACK_MARGINS;
}

/** Only rectangular tables currently expose formatting controls. */
export function textDirectionCellAvailable(state: EditorState): boolean {
	const { $from, $to } = state.selection;
	for (let depth = $from.depth; depth > 0; depth--)
		if ($from.node(depth).type.name === 'table')
			return (
				$from.node(depth).attrs.structureEditable !== false &&
				$to.depth > depth &&
				$to.node(depth) === $from.node(depth)
			);
	return false;
}

/** Only rectangular tables currently expose formatting controls. */
export function tablePropertiesContext(state: EditorState) {
	const { $from, $to } = state.selection;
	for (let depth = $from.depth; depth > 0; depth--) {
		const table = $from.node(depth);
		if (table.type.name !== 'table') continue;
		if (
			table.attrs.structureEditable === false ||
			$to.depth <= depth ||
			$to.node(depth) !== table ||
			$from.depth <= depth
		)
			return null;
		const rows: Array<{ pos: number; properties: TableRowProperties }> = [];
		const cells: Array<{
			pos: number;
			margins: TableCellMargins;
			verticalAlign: TableCell['verticalAlign'];
		}> = [];
		let pos = $from.start(depth);
		table.forEach((row, offset, index) => {
			if (index >= $from.index(depth) && index <= $to.index(depth))
				rows.push({ pos: pos + offset, properties: row.attrs.properties ?? {} });
			row.forEach((cell, cellOffset) => {
				const cellPos = pos + offset + 1 + cellOffset;
				if (
					Math.max(state.selection.from, cellPos + 1) <=
					Math.min(state.selection.to, cellPos + cell.nodeSize - 1)
				)
					cells.push({
						pos: cellPos,
						margins: cell.attrs.margins ? (JSON.parse(cell.attrs.margins) as TableCellMargins) : {},
						verticalAlign: (cell.attrs.verticalAlign ?? undefined) as TableCell['verticalAlign'],
					});
			});
		});
		return {
			pos: $from.before(depth),
			table,
			rows,
			cells,
			margins: (table.attrs.cellMargins
				? JSON.parse(table.attrs.cellMargins)
				: {}) as TableCellMargins,
			inheritedMargins: inheritedCellMargins(table.attrs.inheritedCellMargins),
		};
	}
	return null;
}

/** Patches only fields changed in the dialog; all edits form one history step. */
export function applyTableProperties(
	view: EditorView,
	rowPatch: RowPatch,
	marginPatch: TableCellMargins,
	/** Null removes all direct cell overrides; omitted leaves them untouched. */
	cellMarginPatch?: TableCellMargins | null,
	/** Null removes the direct alignment override; omitted leaves it untouched. */
	cellAlignment?: TableCell['verticalAlign'] | null,
): boolean {
	const context = tablePropertiesContext(view.state);
	if (!view.editable || !context) return false;
	if (
		!Object.keys(rowPatch).length &&
		!Object.keys(marginPatch).length &&
		cellMarginPatch === undefined &&
		cellAlignment === undefined
	)
		return true;
	const tr = view.state.tr;
	for (const row of context.rows) {
		const properties = { ...row.properties, ...rowPatch };
		for (const key of Object.keys(properties) as Array<keyof TableRowProperties>)
			if (properties[key] === undefined) delete properties[key];
		tr.setNodeAttribute(row.pos, 'properties', properties);
		tr.setNodeAttribute(
			row.pos,
			'heightPx',
			properties.heightTwips ? properties.heightTwips / 15 : null,
		);
	}
	if (Object.keys(marginPatch).length)
		tr.setNodeAttribute(
			context.pos,
			'cellMargins',
			JSON.stringify({ ...context.margins, ...marginPatch }),
		);
	if (cellMarginPatch !== undefined || cellAlignment !== undefined)
		for (const cell of context.cells) {
			if (cellMarginPatch !== undefined)
				tr.setNodeAttribute(
					cell.pos,
					'margins',
					cellMarginPatch === null ? null : JSON.stringify({ ...cell.margins, ...cellMarginPatch }),
				);
			if (cellAlignment !== undefined)
				tr.setNodeAttribute(cell.pos, 'verticalAlign', cellAlignment);
		}
	dispatchTableEdit(view, tr);
	return true;
}
