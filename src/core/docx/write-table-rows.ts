// Writes a table whose rows a merge repeated or removed (the rows of a mail merge region): each
// row of the model is a row of the source table, kept, or copied whole with its properties,
// borders, merges and every XML child the model does not hold. Only the cells' paragraphs are
// then written from the model.
import { rowEditKey, type BlockCopies } from './block-copies';
import type { Paragraph, Table } from './model';
import { flatCellParagraphs, flatCells, flatRows } from './parse-content-control';
import { blockChain, placeCellParagraphs, placeControlledById } from './write-block-placement';
import { replaceSlots } from './write-slots';
import { placeTableControl } from './write-table-controls';
import { makeW, type XmlDocument, type XmlElement } from './xml';

type ParagraphWriter = (
	doc: XmlDocument,
	paragraph: Paragraph,
	node: XmlElement,
	base?: Paragraph,
) => XmlElement;

export interface RowCopies {
	copies: BlockCopies;
	/** A deep copy of a source row, for the copy whose first paragraph is `copyId`. */
	clone(source: XmlElement, copyId: string): XmlElement;
}

/** Whether `table` holds rows a merge repeated or removed. */
export const hasMergedRows = (table: Table, rows: RowCopies | undefined): rows is RowCopies =>
	Boolean(rows?.copies.has(rowEditKey(table.id)));

export function writeMergedRows(
	doc: XmlDocument,
	table: Table,
	node: XmlElement,
	base: Table,
	writeParagraph: ParagraphWriter,
	rows: RowCopies,
): XmlElement {
	const flat = flatRows(node);
	const oldRows = flat.map((entry) => entry.element);
	if (oldRows.length !== base.rows.length)
		throw new Error('The rows of a merged table do not match its source.');
	const rowOf = new Map<string, number>();
	base.rows.forEach((row, index) => {
		for (const cell of row) for (const paragraph of cell.paragraphs) rowOf.set(paragraph.id, index);
	});
	const sourceId = (id: string) => (rowOf.has(id) ? id : rows.copies.get(id)?.sourceId);
	const usedRows = new Set<number>();
	const nextRows = table.rows.map((row) => {
		const ids = row.flatMap((cell) => cell.paragraphs.map((paragraph) => paragraph.id));
		const first = ids.find((id) => sourceId(id) !== undefined);
		const source = first === undefined ? undefined : rowOf.get(sourceId(first)!);
		const sourceRow = source === undefined ? undefined : base.rows[source];
		if (source === undefined || !sourceRow || sourceRow.length !== row.length)
			throw new Error('A merged table row has no source row.');
		const copy = !rowOf.has(first!);
		if (!copy && usedRows.has(source)) throw new Error('A source row is written twice.');
		if (!copy) usedRows.add(source);
		const tr = copy ? rows.clone(oldRows[source]!, first!) : oldRows[source]!;
		const cells = flatCells(tr).map((entry) => entry.element);
		row.forEach((cell, index) => {
			const tc = cells[index];
			const baseCell = sourceRow[index]!;
			if (!tc) return;
			const flatParagraphs = flatCellParagraphs(tc);
			const oldParagraphs = flatParagraphs.map((entry) => entry.element);
			const used = new Set<number>();
			const next = cell.paragraphs.map((paragraph) => {
				const id = sourceId(paragraph.id);
				const at = baseCell.paragraphs.findIndex((item) => item.id === id);
				const old = at >= 0 && !used.has(at) ? oldParagraphs[at] : undefined;
				if (old) used.add(at);
				const before = old && { ...baseCell.paragraphs[at]!, id: paragraph.id };
				return writeParagraph(doc, paragraph, old ?? makeW(doc, 'p'), before);
			});
			placeCellParagraphs(doc, tc, flatParagraphs, cell.paragraphs, next);
			// A copied cell control is the copy's own (`placeTableControl` gives it its new id).
			placeTableControl(doc, tc, cell.contentControl, baseCell.contentControl);
		});
		return tr;
	});
	// Each row keeps or joins the row controls the model gives it: a repeating section's new item
	// goes inside the section, after the item before it.
	const chains = table.rows.map((_, index) => blockChain(table.rowProperties?.[index]));
	if (!placeControlledById(doc, node, flat, chains, nextRows, null))
		replaceSlots(node, oldRows, nextRows);
	return node;
}
