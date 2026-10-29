// Table <-> ProseMirror conversion. Merged/nested/complex tables (structureEditable === false)
// preserve their full source row/cell/descriptor structure and only patch paragraph content by
// id, because vertical-merge continuation cells and nested-table previews have no 1:1 visible
// ProseMirror node to rebuild from; simple tables keep the existing rebuild-from-doc path.
import type { Node as ProseMirrorNode } from 'prosemirror-model';
import { cellSourceBlocks, removedByEdit } from 'ooxml-core/docx';
import type {
	ContentControl,
	Paragraph,
	Table,
	TableCell,
	TableStyleCatalog,
} from 'ooxml-core/docx';
import { schema } from './schema';
import { continuationsOf, resolveCellVisuals, rowSpanOf, tableGrid } from 'ooxml-core/docx/ui';
import { convertRestructuredTable, structureChanged } from './table-restructure-convert';

function cellKey(cell: TableCell): string {
	return cell.paragraphs[0]?.id ?? '';
}

export function tableNode(
	block: Table,
	paragraphNode: (paragraph: Paragraph) => ProseMirrorNode,
	tableStyles?: TableStyleCatalog,
	nested = false,
): ProseMirrorNode {
	const columnCount = Math.max(
		1,
		...block.rows.map((row) => row.reduce((sum, cell) => sum + (cell.gridSpan ?? 1), 0)),
	);
	const layout = tableGrid(block);
	const rows = block.rows.map((row, rowIndex) => {
		const cells: ProseMirrorNode[] = [];
		let gridColumn = 0;
		row.forEach((cell, columnIndex) => {
			const column = gridColumn;
			gridColumn += cell.gridSpan ?? 1;
			if (cell.verticalMerge === 'continue') return;
			let rowspan = 1;
			const entry = layout?.rows[rowIndex]?.[columnIndex];
			const continues = layout && entry ? continuationsOf(layout, entry) : [];
			if (layout && entry) rowspan = rowSpanOf(layout, entry);
			else if (cell.verticalMerge === 'restart') {
				let next = rowIndex + 1;
				while (block.rows[next]?.[columnIndex]?.verticalMerge === 'continue') {
					rowspan++;
					next++;
				}
			}
			const visuals = resolveCellVisuals(
				block,
				cell,
				{
					row: rowIndex,
					lastRow: rowIndex + rowspan - 1,
					column,
					lastColumn: column + (cell.gridSpan ?? 1) - 1,
					rowCount: block.rows.length,
					columnCount,
				},
				tableStyles,
			);
			const children = cell.sourceBlocks
				? cellSourceBlocks(cell).map((child) =>
						child.type === 'paragraph'
							? paragraphNode(child)
							: tableNode(child, paragraphNode, tableStyles, true),
					)
				: [
						...cell.paragraphs.map(paragraphNode),
						...(cell.nestedTables ?? []).map((preview) =>
							schema.nodes.nestedTablePreview.create({ rowsJson: JSON.stringify(preview.rows) }),
						),
					];
			cells.push(
				schema.node(
					'tableCell',
					{
						sourceCellKey: cellKey(cell),
						colspan: cell.gridSpan ?? 1,
						rowspan,
						continueKeys: continues.length
							? JSON.stringify(continues.map((c) => cellKey(c.cell)))
							: null,
						widthTwips: cell.widthTwips ?? null,
						verticalAlign: cell.verticalAlign ?? null,
						shadingFill: visuals.shadingFill ?? null,
						directShadingFill: cell.shadingFill ?? null,
						borders: visuals.borders ? JSON.stringify(visuals.borders) : null,
						directBorders: cell.borders ? JSON.stringify(cell.borders) : null,
						margins: cell.margins ? JSON.stringify(cell.margins) : null,
						contentControl: cell.contentControl ?? null,
					},
					children,
				),
			);
		});
		const heightTwips = block.rowProperties?.[rowIndex]?.heightTwips;
		return schema.node(
			'tableRow',
			{
				heightPx: heightTwips ? heightTwips / 15 : null,
				properties: block.rowProperties?.[rowIndex] ?? null,
			},
			cells,
		);
	});
	return schema.node(
		'table',
		{
			id: block.id,
			// Nested tables edit their text on the page; their grid and merges stay as authored.
			structureEditable: !nested && block.structureEditable !== false,
			// Merge and split work on any table whose cells tile its grid.
			restructurable:
				block.restructurable === true || (block.structureEditable !== false && Boolean(layout)),
			grid: block.grid ? JSON.stringify(block.grid) : null,
			widthTwips: block.widthTwips ?? null,
			alignment: block.alignment ?? null,
			indentTwips: block.indentTwips ?? null,
			borders: block.borders ? JSON.stringify(block.borders) : null,
			cellMargins: block.cellMargins ? JSON.stringify(block.cellMargins) : null,
			contentControl: block.contentControl ?? null,
		},
		rows,
	);
}

/**
 * The control around a table or a cell as the editor holds it (picked values, cleared
 * placeholders): a temporary one is gone after an edit. The editor takes no other control off a
 * table or a cell, so one the node lacks (a document shared with an editor that kept none there)
 * stays as `prior` had it, and its `w:sdt` is never unwrapped.
 */
function heldControl(
	node: ProseMirrorNode,
	prior: ContentControl | undefined,
): { contentControl?: ContentControl } {
	const control = (node.attrs.contentControl as ContentControl | null) ?? prior;
	return control && (node.attrs.contentControl || !removedByEdit(control))
		? { contentControl: structuredClone(control) }
		: {};
}

/** The block control around a table as the editor holds it. */
export const tableControl = (node: ProseMirrorNode, prior?: Table): Pick<Table, 'contentControl'> =>
	heldControl(node, prior?.contentControl);

/** Reads editable formatting from nodes rather than a stale prior model. */
export function tableFormattingFromNode(node: ProseMirrorNode): {
	rowProperties: Table['rowProperties'];
	cellMargins: Table['cellMargins'];
} {
	const rowProperties: NonNullable<Table['rowProperties']> = [];
	node.forEach((row) => rowProperties.push(structuredClone(row.attrs.properties ?? {})));
	return {
		rowProperties: rowProperties.some((row) => Object.keys(row).length) ? rowProperties : undefined,
		cellMargins:
			typeof node.attrs.cellMargins === 'string'
				? (JSON.parse(node.attrs.cellMargins) as Table['cellMargins'])
				: undefined,
	};
}

/** Rebuilds a simple (structurally editable) table directly from its visible ProseMirror rows. */
export function convertSimpleTable(
	node: ProseMirrorNode,
	convertParagraph: (paragraph: ProseMirrorNode) => Paragraph,
	prior?: Table,
): Table['rows'] {
	// Preserve untouched descriptors by source cell identity; direct border, shading and margin edits use
	// node attributes so editing text never flattens inherited table-style formatting.
	const priorCells = new Map<string, TableCell>();
	for (const row of prior?.rows ?? []) for (const cell of row) priorCells.set(cellKey(cell), cell);
	const rows: Table['rows'] = [];
	node.forEach((row) => {
		const cells: Table['rows'][number] = [];
		row.forEach((cell) => {
			const paragraphs: Paragraph[] = [];
			cell.forEach((paragraph) => {
				if (paragraph.type.name === 'paragraph') paragraphs.push(convertParagraph(paragraph));
			});
			const source = priorCells.get(String(cell.attrs.sourceCellKey || ''));
			const {
				paragraphs: _previous,
				borders: _borders,
				margins: _margins,
				verticalAlign: _verticalAlign,
				contentControl: _control,
				...properties
			} = source ?? { paragraphs: [] };
			const direct = typeof cell.attrs.directBorders === 'string' ? cell.attrs.directBorders : null;
			cells.push({
				...structuredClone(properties),
				...(direct ? { borders: JSON.parse(direct) as NonNullable<TableCell['borders']> } : {}),
				...(typeof cell.attrs.verticalAlign === 'string'
					? { verticalAlign: cell.attrs.verticalAlign as NonNullable<TableCell['verticalAlign']> }
					: {}),
				...(typeof cell.attrs.margins === 'string'
					? { margins: JSON.parse(cell.attrs.margins) as NonNullable<TableCell['margins']> }
					: {}),
				...heldControl(cell, source?.contentControl),
				paragraphs,
			});
			if (cell.attrs.shadingEdited) {
				const target = cells[cells.length - 1]!;
				delete target.shadingThemeFill;
				target.shadingFill = String(cell.attrs.directShadingFill);
			}
		});
		rows.push(cells);
	});
	return rows;
}

/**
 * Patches paragraph content back into a merged/nested/complex table's prior structure by cell
 * identity, leaving row/cell layout, merges and descriptive properties exactly as they were.
 */
export function convertMergedTable(
	node: ProseMirrorNode,
	prior: Table,
	convertParagraph: (paragraph: ProseMirrorNode) => Paragraph,
): Table {
	const updated = structuredClone(prior);
	delete updated.contentControl;
	Object.assign(updated, tableControl(node, prior));
	const cellByKey = new Map<string, TableCell>();
	for (const row of updated.rows) for (const cell of row) cellByKey.set(cellKey(cell), cell);
	// A row's control as the editor holds it (a temporary one goes after an edit); the rows of a
	// merged table stay as they were, one node per row.
	node.forEach((row, _offset, index) => {
		const properties = updated.rowProperties?.[index];
		if (!properties) return;
		delete properties.contentControl;
		const own = row.attrs.properties as NonNullable<Table['rowProperties']>[number] | null;
		const control = own?.contentControl;
		if (control) properties.contentControl = structuredClone(control);
	});
	node.forEach((row) =>
		row.forEach((cellNode) => {
			const target = cellByKey.get(String(cellNode.attrs.sourceCellKey || ''));
			if (!target) return;
			const paragraphs: Paragraph[] = [];
			const nested: ProseMirrorNode[] = [];
			cellNode.forEach((child) => {
				if (child.type.name === 'paragraph') paragraphs.push(convertParagraph(child));
				else if (child.type.name === 'table') nested.push(child);
			});
			if (paragraphs.length) target.paragraphs = paragraphs;
			const control = target.contentControl;
			delete target.contentControl;
			Object.assign(target, heldControl(cellNode, control));
			let index = 0;
			if (target.sourceBlocks)
				target.sourceBlocks = target.sourceBlocks.map((block) => {
					const node = block.type === 'table' ? nested[index++] : undefined;
					return node && block.type === 'table'
						? convertStructuredTable(node, block, convertParagraph)
						: block;
				});
		}),
	);
	return updated;
}

/**
 * A merged or nested table back to the model: rebuilt from the node when its cells, merges or
 * grid were edited, otherwise patched by cell identity so its structure stays as authored.
 */
export function convertStructuredTable(
	node: ProseMirrorNode,
	prior: Table,
	convertParagraph: (paragraph: ProseMirrorNode) => Paragraph,
): Table {
	if (!node.attrs.restructurable || !structureChanged(node, prior))
		return convertMergedTable(node, prior, convertParagraph);
	const updated = convertRestructuredTable(
		node,
		prior,
		convertParagraph,
		(inner, innerPrior) => convertStructuredTable(inner, innerPrior, convertParagraph),
		heldControl,
	);
	delete updated.contentControl;
	Object.assign(updated, tableControl(node, prior));
	return updated;
}
