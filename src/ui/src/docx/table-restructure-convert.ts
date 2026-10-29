// Editor table node -> core model for a merged or nested table whose structure was edited. Every
// surviving cell starts from its source cell (so the writer keeps its element and any property the
// model does not hold); spans, merges, widths and the grid come from the edited node.
import type { ContentControl, Paragraph, Table, TableCell } from 'ooxml-core/docx';
import type { Node as ProseMirrorNode } from 'prosemirror-model';
import { rowSpanOf, tableGrid } from 'ooxml-core/docx/ui';
import { gridOf, slotsOf } from './table-restructure';

const keyOf = (cell: TableCell): string => cell.paragraphs[0]?.id ?? '';

/** Whether the node's cells, spans, merges, widths or grid differ from the model's. */
export function structureChanged(node: ProseMirrorNode, prior: Table): boolean {
	const fromNode = slotsOf(node).map((row) =>
		row.map((slot) => [
			slot.key,
			slot.span,
			slot.offset ? 'continue' : Number(slot.cell.attrs.rowspan) > 1 ? 'restart' : '',
			slot.cell.attrs.widthTwips ?? null,
		]),
	);
	const layout = tableGrid(prior);
	const fromModel = prior.rows.map((row, r) =>
		row.map((cell, index) => {
			const entry = layout?.rows[r]?.[index];
			// A merge start with nothing merged into it is an ordinary cell as far as the editor sees.
			const merge =
				cell.verticalMerge === 'restart' && layout && entry && rowSpanOf(layout, entry) === 1
					? ''
					: (cell.verticalMerge ?? '');
			return [keyOf(cell), cell.gridSpan ?? 1, merge, cell.widthTwips ?? null];
		}),
	);
	return (
		JSON.stringify(fromNode) !== JSON.stringify(fromModel) ||
		JSON.stringify(gridOf(node) ?? null) !== JSON.stringify(prior.grid ?? null)
	);
}

export type NestedConverter = (node: ProseMirrorNode, prior: Table) => Table;

/**
 * Rebuilds `prior` from the edited `node`: rows, cells and merges as the node holds them, each
 * surviving cell cloned from its source cell by first-paragraph id.
 */
export function convertRestructuredTable(
	node: ProseMirrorNode,
	prior: Table,
	convertParagraph: (paragraph: ProseMirrorNode) => Paragraph,
	convertNested: NestedConverter,
	heldControl: (cell: ProseMirrorNode, prior: ContentControl | undefined) => Partial<TableCell>,
): Table {
	const updated = structuredClone(prior);
	// From here the table is saved through the merge-aware writer, whatever its source shape.
	updated.structureEditable = false;
	updated.restructurable = true;
	const sources = new Map<string, TableCell>();
	for (const row of prior.rows) for (const cell of row) sources.set(keyOf(cell), cell);
	const before = tableGrid(prior);
	const merged = new Map<string, number>();
	for (const entry of before?.rows.flat() ?? [])
		if (entry.cell.verticalMerge === 'restart')
			merged.set(keyOf(entry.cell), rowSpanOf(before!, entry));
	updated.rows = slotsOf(node).map((slots) =>
		slots.map((slot): TableCell => {
			const base = sources.get(slot.key);
			const cell: TableCell = base ? structuredClone(base) : { paragraphs: [] };
			if (slot.span > 1) cell.gridSpan = slot.span;
			else delete cell.gridSpan;
			const width = slot.cell.attrs.widthTwips as TableCell['widthTwips'] | null;
			if (width != null) cell.widthTwips = width;
			else delete cell.widthTwips;
			if (slot.offset > 0) {
				cell.verticalMerge = 'continue';
				if (!cell.paragraphs.length)
					cell.paragraphs = [{ type: 'paragraph', id: slot.key, runs: [{ text: '' }] }];
				return cell;
			}
			if (Number(slot.cell.attrs.rowspan) > 1) cell.verticalMerge = 'restart';
			// A merge start with nothing merged into it that the source already had stays as it was.
			else if (!(base?.verticalMerge === 'restart' && (merged.get(slot.key) ?? 1) === 1))
				delete cell.verticalMerge;
			const paragraphs: Paragraph[] = [];
			const nested: ProseMirrorNode[] = [];
			slot.cell.forEach((child) => {
				if (child.type.name === 'paragraph') paragraphs.push(convertParagraph(child));
				else if (child.type.name === 'table') nested.push(child);
			});
			if (paragraphs.length) cell.paragraphs = paragraphs;
			const attrs = slot.cell.attrs;
			if (!base || attrs.shadingEdited) {
				delete cell.shadingThemeFill;
				if (attrs.directShadingFill) cell.shadingFill = String(attrs.directShadingFill);
				else delete cell.shadingFill;
			}
			const borders = typeof attrs.directBorders === 'string' ? attrs.directBorders : null;
			if (borders) cell.borders = JSON.parse(borders) as NonNullable<TableCell['borders']>;
			else if (!base) delete cell.borders;
			if (typeof attrs.margins === 'string')
				cell.margins = JSON.parse(attrs.margins) as NonNullable<TableCell['margins']>;
			if (typeof attrs.verticalAlign === 'string')
				cell.verticalAlign = attrs.verticalAlign as NonNullable<TableCell['verticalAlign']>;
			delete cell.contentControl;
			Object.assign(cell, heldControl(slot.cell, base?.contentControl));
			let index = 0;
			if (cell.sourceBlocks)
				cell.sourceBlocks = cell.sourceBlocks.map((block) => {
					const target = block.type === 'table' ? nested[index++] : undefined;
					return target && block.type === 'table' ? convertNested(target, block) : block;
				});
			return cell;
		}),
	);
	const grid = gridOf(node);
	if (grid) updated.grid = grid as NonNullable<Table['grid']>;
	const properties: NonNullable<Table['rowProperties']> = [];
	node.forEach((row) => properties.push(structuredClone(row.attrs.properties ?? {})));
	if (properties.some((entry) => Object.keys(entry).length)) updated.rowProperties = properties;
	else delete updated.rowProperties;
	return updated;
}
