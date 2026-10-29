// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Saves a restructured merged or nested table (rows or columns added or removed, cells merged or
// split) by patching the source `w:tbl`: a row or cell that survives keeps its element, so every
// property the model does not hold (unknown `tblPr`, `trPr` and `tcPr` children, extension
// attributes) stays; only the grid, spans, widths, merges and the cells the edit created change.
import { flatCellParagraphs, flatCells, flatRows } from './parse-content-control';
import type { Paragraph, Table, TableCell } from './model';
import { buildCellProperties } from './table-cell-write';
import {
	CELL_ORDER,
	patchCellBorders,
	patchCellMargins,
	patchCellShading,
	patchCellVerticalAlignment,
	patchRowProperties,
	patchTableMargins,
} from './write-table-properties';
import { orderChildren } from './element-order';
import { children, first, makeW, WORD_NS, type XmlDocument, type XmlElement } from './xml';

export type ParagraphWriter = (
	doc: XmlDocument,
	paragraph: Paragraph,
	node: XmlElement,
	base?: Paragraph,
) => XmlElement;
export type SlotsWriter = (
	doc: XmlDocument,
	parent: XmlElement,
	old: XmlElement[],
	next: XmlElement[],
) => void;
/** Writes the tables nested in a surviving cell back into that cell's source XML. */
export type NestedWriter = (
	doc: XmlDocument,
	tc: XmlElement,
	cell: TableCell,
	base: TableCell,
) => void;

const set = (element: XmlElement, name: string, value: string): void =>
	element.setAttributeNS(WORD_NS, `w:${name}`, value);

/** True when `next` lays the table out differently from `base`: another grid, span or merge. */
export function restructured(next: Table, base: Table): boolean {
	const shape = (table: Table) =>
		JSON.stringify([
			table.grid,
			table.rows.map((row) =>
				row.map((cell) => [
					cell.paragraphs.map((p) => p.id),
					cell.gridSpan ?? 1,
					cell.verticalMerge ?? '',
					cell.widthTwips,
				]),
			),
		]);
	return shape(next) !== shape(base);
}

/** Sets or removes `tcW`, `gridSpan` and `vMerge` on a surviving cell, keeping every other child. */
function patchCellLayout(doc: XmlDocument, tc: XmlElement, cell: TableCell, base: TableCell): void {
	const span = cell.gridSpan ?? 1;
	const width = cell.widthTwips;
	if (
		span === (base.gridSpan ?? 1) &&
		width === base.widthTwips &&
		cell.verticalMerge === base.verticalMerge
	)
		return;
	let props = first(tc, 'tcPr');
	if (!props) {
		props = makeW(doc, 'tcPr');
		tc.insertBefore(props, tc.firstChild);
	}
	const replace = (name: string, next: XmlElement | undefined) => {
		for (const old of children(props, name)) props.removeChild(old);
		if (next) props.appendChild(next);
	};
	if (width !== base.widthTwips) {
		const old = first(props, 'tcW');
		if (width === undefined) replace('tcW', undefined);
		else {
			const element = old ?? makeW(doc, 'tcW');
			set(element, 'w', String(width));
			set(element, 'type', 'dxa');
			if (!old) props.appendChild(element);
		}
	}
	if (span !== (base.gridSpan ?? 1)) {
		const element = span > 1 ? makeW(doc, 'gridSpan') : undefined;
		if (element) set(element, 'val', String(span));
		replace('gridSpan', element);
	}
	if (cell.verticalMerge !== base.verticalMerge) {
		const element = cell.verticalMerge ? makeW(doc, 'vMerge') : undefined;
		if (element && cell.verticalMerge === 'restart') set(element, 'val', 'restart');
		replace('vMerge', element);
	}
	orderChildren(props, CELL_ORDER);
	if (!Array.from(props.childNodes).some((node) => node.nodeType === 1) && !props.attributes.length)
		tc.removeChild(props);
}

/** Brings `w:tblGrid` to the model's column widths, keeping the `gridCol` elements that remain. */
function patchGrid(doc: XmlDocument, node: XmlElement, table: Table): void {
	const grid = first(node, 'tblGrid');
	if (!grid || !table.grid) return;
	const old = children(grid, 'gridCol');
	table.grid.forEach((width, index) => {
		let column = old[index];
		if (!column) {
			column = makeW(doc, 'gridCol');
			grid.appendChild(column);
		}
		if (column.getAttributeNS(WORD_NS, 'w') !== String(width)) set(column, 'w', String(width));
	});
	for (const extra of old.slice(table.grid.length)) grid.removeChild(extra);
}

/**
 * Applies the restructured `table` to its source `node`. `base` is the table as it was loaded; the
 * caller has checked that the source can be restructured (`canRestructureTable`).
 */
export function writeRestructuredTable(
	doc: XmlDocument,
	table: Table,
	node: XmlElement,
	base: Table,
	writeParagraph: ParagraphWriter,
	replaceSlots: SlotsWriter,
	writeNested: NestedWriter,
): XmlElement {
	if (table.rows.some((row) => !row.length || row.some((cell) => !cell.paragraphs.length)))
		throw new Error('Every table row needs a cell and every cell a paragraph.');
	const oldRows = flatRows(node).map((entry) => entry.element);
	const oldCells = oldRows.map((tr) => flatCells(tr).map((entry) => entry.element));
	const cellById = new Map<string, { tc: XmlElement; model: TableCell; row: number }>();
	const paragraphById = new Map<string, { element: XmlElement; model: Paragraph }>();
	base.rows.forEach((row, ri) =>
		row.forEach((cell, ci) => {
			const tc = oldCells[ri]?.[ci];
			if (!tc) return;
			const id = cell.paragraphs[0]?.id;
			if (id) cellById.set(id, { tc, model: cell, row: ri });
			const flat = flatCellParagraphs(tc);
			cell.paragraphs.forEach((model, pi) => {
				const entry = flat[pi];
				if (entry) paragraphById.set(model.id, { element: entry.element, model });
			});
		}),
	);
	const usedRows = new Set<number>();
	const usedCells = new Set<XmlElement>();
	const nextRows = table.rows.map((row, rowIndex) => {
		const sourceRow = row
			.map((cell) => cellById.get(cell.paragraphs[0]?.id ?? '')?.row)
			.find((index): index is number => index !== undefined && !usedRows.has(index));
		if (sourceRow !== undefined) usedRows.add(sourceRow);
		const tr =
			sourceRow === undefined ? makeW(doc, 'tr') : (oldRows[sourceRow] ?? makeW(doc, 'tr'));
		patchRowProperties(
			doc,
			tr,
			table.rowProperties?.[rowIndex],
			sourceRow === undefined ? undefined : base.rowProperties?.[sourceRow],
		);
		const nextCells = row.map((cell) => {
			const source = cellById.get(cell.paragraphs[0]?.id ?? '');
			const reuse =
				source && source.row === sourceRow && !usedCells.has(source.tc) ? source : undefined;
			const tc = reuse?.tc ?? makeW(doc, 'tc');
			if (reuse) {
				usedCells.add(reuse.tc);
				patchCellLayout(doc, tc, cell, reuse.model);
				patchCellBorders(doc, tc, cell.borders, reuse.model.borders);
				patchCellShading(doc, tc, cell, reuse.model);
				patchCellMargins(doc, tc, cell.margins, reuse.model.margins);
				patchCellVerticalAlignment(doc, tc, cell.verticalAlign, reuse.model.verticalAlign);
			} else {
				const props = buildCellProperties(doc, cell);
				if (props) tc.insertBefore(props, tc.firstChild);
			}
			const oldParagraphs = reuse ? flatCellParagraphs(tc).map((entry) => entry.element) : [];
			const owned = new Set(oldParagraphs);
			// A paragraph that changed cells (a merge) moves as a copy: its source cell is dropped or
			// emptied by the same edit.
			const nextParagraphs = cell.paragraphs.map((paragraph) => {
				const found = paragraphById.get(paragraph.id);
				const element = found
					? owned.has(found.element)
						? found.element
						: (found.element.cloneNode(true) as XmlElement)
					: makeW(doc, 'p');
				return writeParagraph(doc, paragraph, element, found?.model);
			});
			replaceSlots(doc, tc, oldParagraphs, nextParagraphs);
			if (reuse) writeNested(doc, tc, cell, reuse.model);
			return tc;
		});
		replaceSlots(doc, tr, sourceRow === undefined ? [] : (oldCells[sourceRow] ?? []), nextCells);
		return tr;
	});
	replaceSlots(doc, node, oldRows, nextRows);
	patchTableMargins(doc, table, node, base);
	patchGrid(doc, node, table);
	return node;
}
