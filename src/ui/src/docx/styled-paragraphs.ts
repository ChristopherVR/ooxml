import type { Node as ProseMirrorNode } from 'prosemirror-model';
import {
	resolveTableStyleFormatting,
	type DocumentModel,
	type TableStyleConditionalFormatting,
} from 'ooxml-core/docx';
import { gridOf, slotsOf } from './table-restructure';

const grids = new WeakMap<ProseMirrorNode, { columns: number; starts: number[][] }>();

function tableGeometry(table: ProseMirrorNode) {
	const cached = grids.get(table);
	if (cached) return cached;
	const slots = slotsOf(table);
	const columns =
		gridOf(table)?.length ??
		Math.max(
			1,
			...slots.map((row) => {
				const last = row.at(-1);
				return last ? last.col + last.span : 0;
			}),
		);
	const starts = slots.map((row) =>
		row.filter((slot) => slot.offset === 0).map((slot) => slot.col),
	);
	const geometry = { columns, starts };
	grids.set(table, geometry);
	return geometry;
}

function cellFormatting(table: ProseMirrorNode, model: DocumentModel) {
	const prior = table.attrs.tablePropertiesKnown
		? undefined
		: model.blocks.find((block) => block.type === 'table' && block.id === table.attrs.id);
	const styleId = table.attrs.tableStyleId ?? (prior?.type === 'table' ? prior.style : undefined);
	if (!styleId || !model.tableStyles) return;
	const look = table.attrs.tableLook ?? (prior?.type === 'table' ? prior.look : undefined);
	const { columns, starts } = tableGeometry(table);
	return (row: number, cell: number) =>
		resolveTableStyleFormatting(
			styleId,
			model.tableStyles,
			look,
			row,
			table.childCount,
			starts[row]?.[cell] ?? 0,
			columns,
		);
}

/** The nearest table cell's formatting at a paragraph boundary or a text cursor. */
export function tableStyleAt(doc: ProseMirrorNode, pos: number, model: DocumentModel | undefined) {
	if (!model?.tableStyles) return;
	const $pos = doc.resolve(pos);
	for (let depth = $pos.depth - 2; depth > 0; depth--) {
		const table = $pos.node(depth);
		if (table.type.name === 'table')
			return cellFormatting(table, model)?.($pos.index(depth), $pos.index(depth + 1));
	}
}

/** Walks visible ProseMirror paragraphs with their current table cell formatting and node position. */
export function visitStyledParagraphs(
	doc: ProseMirrorNode,
	model: DocumentModel,
	visit: (node: ProseMirrorNode, pos: number, tableStyle?: TableStyleConditionalFormatting) => void,
): void {
	const walk = (
		node: ProseMirrorNode,
		pos: number,
		tableStyle?: TableStyleConditionalFormatting,
	) => {
		if (node.type.name === 'paragraph') {
			visit(node, pos, tableStyle);
			return;
		}
		if (node.type.name !== 'table' || !model.tableStyles) {
			node.forEach((child, offset) => walk(child, pos + 1 + offset, tableStyle));
			return;
		}

		const resolve = cellFormatting(node, model);
		node.forEach((row, rowOffset, rowIndex) => {
			row.forEach((cell, cellOffset, cellIndex) => {
				const formatting = resolve?.(rowIndex, cellIndex);
				const start = pos + rowOffset + cellOffset + 3;
				cell.forEach((child, offset) => walk(child, start + offset, formatting));
			});
		});
	};
	doc.forEach((node, pos) => walk(node, pos));
}
