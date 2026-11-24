/**
 * List structure for the tagged Word PDF export (ISO 32000-1 14.8.4.3.3, 14.8.5.5): numbered
 * and bulleted paragraphs become `L` > `LI` > `Lbl` (the marker) + `LBody` (the text), items of
 * one list grouped in one `L`, deeper levels as an `L` nested in the `LBody` of the item above,
 * each `L` saying how it is numbered (`/ListNumbering`). DOM-free.
 */
import {
	computeListLabels,
	resolveNumberingLevel,
	type DocumentModel,
	type ParagraphListLabel,
} from 'ooxml-core/docx';
import type { PdfStructElement } from '../../export/pdf-vector';

/** The `ListNumbering` values of ISO 32000-1 Table 347. */
export type PdfListNumbering =
	| 'None'
	| 'Disc'
	| 'Circle'
	| 'Square'
	| 'Decimal'
	| 'UpperRoman'
	| 'LowerRoman'
	| 'UpperAlpha'
	| 'LowerAlpha';

/** The list a paragraph is an item of. */
export interface DocxPdfListItem {
	numId: string;
	/** 0-based list level. */
	level: number;
	numbering: PdfListNumbering;
}

const CIRCLES = new Set(['o', '\u25cb', '\u25e6', '\u26aa']);
const SQUARES = new Set(['\uf0a7', '\uf06e', '\uf0a8', '\u25aa', '\u25a0', '\u25a1', '\u25ab']);

/** How a level's markers read: its number format, or the shape of its bullet. */
export function listNumbering(numFmt: string, marker: string): PdfListNumbering {
	switch (numFmt) {
		case 'bullet': {
			const glyph = marker.trim();
			if (CIRCLES.has(glyph)) return 'Circle';
			if (SQUARES.has(glyph)) return 'Square';
			// Word's Symbol bullet and every other bullet glyph read as a bullet.
			return 'Disc';
		}
		case 'upperRoman':
			return 'UpperRoman';
		case 'lowerRoman':
			return 'LowerRoman';
		case 'upperLetter':
			return 'UpperAlpha';
		case 'lowerLetter':
			return 'LowerAlpha';
		case 'none':
			return 'None';
		default:
			// Decimal, ordinals, number words and other scripts' counting are ordered numbers.
			return 'Decimal';
	}
}

/** The list item of every numbered or bulleted paragraph of the body, by paragraph id. */
export function documentListItems(model: DocumentModel): Map<string, DocxPdfListItem> {
	const items = new Map<string, DocxPdfListItem>();
	const catalog = model.numberingCatalog;
	if (!catalog) return items;
	const labels: Map<string, ParagraphListLabel> = computeListLabels(model);
	for (const [id, label] of labels) {
		const numFmt = resolveNumberingLevel(catalog, label.numId, label.level)?.numFmt ?? 'decimal';
		const numbering = label.picture ? 'Disc' : listNumbering(numFmt, label.text);
		items.set(id, { numId: label.numId, level: label.level, numbering });
	}
	return items;
}

interface OpenList {
	item: DocxPdfListItem;
	list: PdfStructElement;
	body?: PdfStructElement;
}

/** The lists open in one container (the document or a table cell), nested by level. */
export class DocxPdfListNesting {
	readonly #open: OpenList[] = [];

	/** Anything else placed in the container ends its lists. */
	reset(): void {
		this.#open.length = 0;
	}

	/** Adds an item to its list (opening the list, nested under the item above when deeper). */
	add(
		container: PdfStructElement,
		item: DocxPdfListItem,
	): { label: PdfStructElement; body: PdfStructElement } {
		const open = this.#open;
		while (open.length && open.at(-1)!.item.level > item.level) open.pop();
		const same = open.at(-1);
		if (
			same &&
			same.item.level === item.level &&
			(same.item.numId !== item.numId || same.item.numbering !== item.numbering)
		)
			open.pop();
		let top = open.at(-1);
		if (!top || top.item.level < item.level) {
			const list: PdfStructElement = {
				type: 'L',
				kids: [],
				attributes: { List: { ListNumbering: item.numbering } },
			};
			(top?.body ?? container).kids.push(list);
			top = { item, list };
			open.push(top);
		}
		const label: PdfStructElement = { type: 'Lbl', kids: [] };
		const body: PdfStructElement = { type: 'LBody', kids: [] };
		top.list.kids.push({ type: 'LI', kids: [label, body] });
		top.body = body;
		return { label, body };
	}
}
