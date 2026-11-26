// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import type { Block, DocumentModel, Paragraph, SectionProperties } from './model';
import { copySourceBlock, type BlockCopies } from './block-copies';
import { cellSourceBlocks } from './cell-blocks';
import { flatBlocks } from './parse-content-control';
import { placeBlocksById, wrapNewBlockControls } from './write-block-placement';
import { replaceSlots } from './write-slots';
import { children, first, makeW, type XmlDocument, type XmlElement, WORD_NS } from './xml';
import { commentContinuations } from './comment-spans';
import { annotationIdAllocator, newMoveRangeIds, type CommentSpans } from './write-ranges';
import { writeTable as writeTableContent } from './write-table';
import { placeTableControl } from './write-table-controls';
import { buildNewTableProperties, newTableGridWidths, tableColumnCount } from './table-defaults';
import { twips as twipsOf } from './units';
import { buildCellProperties, buildRowProperties } from './table-cell-write';
import { orderSectionProperties } from './element-order';
import { applySectionEdits } from './write-sections';
import { createParagraph, setAttribute, writeParagraphImpl } from './write-paragraph';
import type { DocPrIdAllocator } from './docpr-ids';
import {
	RelationshipAllocator,
	scanUsedRelationshipIds,
	type NewRelationship,
} from './relationship-allocator';

const twips = (pixels: number): string => String(Math.round(pixels * 15));
function createTable(
	doc: XmlDocument,
	table: Extract<Block, { type: 'table' }>,
	allocator: RelationshipAllocator | undefined,
	contentWidthTwips: number,
	spans?: CommentSpans,
): XmlElement {
	const node = makeW(doc, 'tbl');
	const { tblPr, tblGrid } = buildNewTableProperties(doc, table, contentWidthTwips);
	node.appendChild(tblPr);
	node.appendChild(tblGrid);
	const columns = tableColumnCount(table);
	const grid = newTableGridWidths(table, contentWidthTwips);
	table.rows.forEach((row, rowIndex) => {
		const tr = makeW(doc, 'tr');
		const spanned = row.reduce((sum, cell) => sum + (cell.gridSpan ?? 1), 0);
		const trPr = buildRowProperties(doc, table.rowProperties?.[rowIndex], columns - spanned);
		if (trPr) tr.appendChild(trPr);
		let start = 0;
		for (const cell of row) {
			const tc = makeW(doc, 'tc');
			// Word sizes an AutoFit table without cell widths from its content, ignoring the grid:
			// give each new cell its grid width as its preferred width, as Word's Insert Table does.
			const span = cell.gridSpan ?? 1;
			const gridWidth = grid.slice(start, start + span).reduce((sum, width) => sum + width, 0);
			start += span;
			const tcPr = buildCellProperties(
				doc,
				cell.widthTwips === undefined && cell.widthPercent === undefined && gridWidth > 0
					? { ...cell, widthTwips: twipsOf(gridWidth) }
					: cell,
			);
			if (tcPr) tc.appendChild(tcPr);
			for (const block of cellSourceBlocks(cell))
				tc.appendChild(
					block.type === 'paragraph'
						? createParagraph(doc, block, allocator, spans)
						: createTable(doc, block, allocator, Math.max(1, gridWidth), spans),
				);
			if (!cell.paragraphs.length) tc.appendChild(makeW(doc, 'p'));
			tr.appendChild(tc);
			placeTableControl(doc, tc, cell.contentControl, undefined);
		}
		node.appendChild(tr);
		placeTableControl(doc, tr, table.rowProperties?.[rowIndex]?.contentControl, undefined);
	});
	return node;
}

function baseMap(blocks: Block[]): Map<string, Block> {
	return new Map(blocks.map((block) => [block.id, block]));
}

/**
 * Applies `blocks` to the paragraphs and tables directly inside `container` (the document body, or
 * a header, footer or note root), rewriting only blocks that changed relative to `original`.
 * Without an allocator, new pictures and external links are rejected (their relationships would
 * belong to a different package part).
 */
export function applyBlocks(
	doc: XmlDocument,
	container: XmlElement,
	blocks: Block[],
	original: Block[],
	allocator: RelationshipAllocator | undefined,
	contentWidthTwips: number,
	anchor: XmlElement | null = null,
	copies?: BlockCopies,
): void {
	const nextId = annotationIdAllocator(doc, blocks);
	const spans: CommentSpans = {
		next: commentContinuations(blocks),
		original: commentContinuations(original),
		rangeIds: newMoveRangeIds(blocks, nextId),
		nextId,
	};
	const boundWriteParagraph = (
		writeDoc: XmlDocument,
		paragraph: Paragraph,
		node: XmlElement,
		base?: Paragraph,
	) => writeParagraphImpl(writeDoc, paragraph, node, base, allocator, spans);
	const flat = flatBlocks(container);
	const oldNodes = flat.map((entry) => entry.element);
	const oldById = new Map<string, XmlElement>();
	original.forEach((block, index) => {
		if (oldNodes[index]) oldById.set(block.id, oldNodes[index]);
	});
	const bases = baseMap(original);
	const copied = new Map<string, ReturnType<typeof copySourceBlock>>();
	const bookmarks = new Map<string, string>();
	for (const [id, copy] of copies ?? []) {
		const source = oldById.get(copy.sourceId);
		const base = bases.get(copy.sourceId);
		if (!source || !base) continue;
		copied.set(
			id,
			copySourceBlock(source, base, copy, allocator, (key) => {
				const known = bookmarks.get(key);
				if (known) return known;
				const value = nextId();
				bookmarks.set(key, value);
				return value;
			}),
		);
	}
	const output: XmlElement[] = [];
	for (const block of blocks) {
		const copy = copied.get(block.id);
		const old = oldById.get(block.id) ?? copy?.node;
		const base = bases.get(block.id) ?? copy?.base;
		if (block.type === 'paragraph')
			output.push(
				old
					? boundWriteParagraph(doc, block, old, base?.type === 'paragraph' ? base : undefined)
					: createParagraph(doc, block, allocator, spans),
			);
		else
			output.push(
				old
					? writeTableContent(
							doc,
							block,
							old,
							base?.type === 'table' ? base : undefined,
							boundWriteParagraph,
						)
					: createTable(doc, block, allocator, contentWidthTwips, spans),
			);
	}
	const nested = flat.some((entry) => entry.sdt) || blocks.some((block) => block.contentControl);
	if (!nested || !placeBlocksById(doc, container, flat, blocks, output, anchor))
		replaceSlots(container, oldNodes, output, anchor);
	if (nested) wrapNewBlockControls(doc, blocks, output);
}

export interface ApplyModelResult {
	/** New hyperlink/image relationships save.ts must add to word/_rels/document.xml.rels. */
	newRelationships: readonly NewRelationship[];
}

export function applyModel(
	doc: XmlDocument,
	model: DocumentModel,
	original: Block[],
	/** Relationship ids already declared in word/_rels/document.xml.rels (styles, numbering, ...). */
	reservedRelationshipIds: Iterable<string> = [],
	/** The loaded document's sections, so unchanged sections stay byte-identical. */
	baseSections: SectionProperties[] = [],
	/** Package-wide `wp:docPr` id allocator shared with header, footer and note parts. */
	docPrIds?: DocPrIdAllocator,
	/** Relationship ids of header/footer parts created for this save, by part name. */
	headerFooterReferences: ReadonlyMap<string, string> = new Map(),
	copies?: BlockCopies,
): ApplyModelResult {
	const body = Array.from(doc.getElementsByTagNameNS(WORD_NS, 'body'))[0];
	if (!body) throw new Error('DOCX document.xml has no w:body');
	const allocator = new RelationshipAllocator(
		[...scanUsedRelationshipIds(doc), ...reservedRelationshipIds],
		docPrIds,
	);
	allocator.docPrIds.reserveFromDocument(doc);
	const contentWidthTwips = Math.round(
		(model.page.width - model.page.marginLeft - model.page.marginRight) * 15,
	);
	applyBlocks(
		doc,
		body,
		model.blocks,
		original,
		allocator,
		contentWidthTwips,
		children(body, 'sectPr').at(-1) ?? null,
		copies,
	);
	let section = children(body, 'sectPr').at(-1);
	if (!section) {
		section = makeW(doc, 'sectPr');
		body.appendChild(section);
	}
	let size = first(section, 'pgSz');
	if (!size) {
		size = makeW(doc, 'pgSz');
		section.appendChild(size);
	}
	setAttribute(size, 'w', twips(model.page.width));
	setAttribute(size, 'h', twips(model.page.height));
	let margins = first(section, 'pgMar');
	if (!margins) {
		margins = makeW(doc, 'pgMar');
		section.appendChild(margins);
	}
	for (const side of ['Top', 'Right', 'Bottom', 'Left'] as const)
		setAttribute(margins, side.toLowerCase(), twips(model.page[`margin${side}`]));
	// CT_PageMar requires header, footer and gutter; keep a source's values, else Word's defaults.
	for (const [name, fallback] of [
		['header', '720'],
		['footer', '720'],
		['gutter', '0'],
	] as const)
		if (!margins.getAttributeNS(WORD_NS, name)) setAttribute(margins, name, fallback);
	orderSectionProperties(section);
	if (model.sections)
		applySectionEdits(
			doc,
			body,
			model.blocks,
			model.sections,
			baseSections,
			headerFooterReferences,
		);
	return { newRelationships: allocator.newRelationships };
}
