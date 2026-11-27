// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// The paragraph-mark half of the paragraph writer: `w:pPr` from the model. Kept apart from the
// inline writer (`write-paragraph.ts`) so a properties-only caller, such as the formatting
// revision snapshot of the Word editor, does not reach the run, drawing and body writers.
import type { Paragraph } from './model';
import {
	children,
	first,
	importElement,
	makeW,
	WORD_NS,
	type XmlDocument,
	type XmlElement,
} from './xml';
import { writeParagraphProperties } from './write-paragraph-properties';
import { orderParagraphProperties, writeTabStops } from './tab-stops';
import { writeNumberingProperties } from './numbering-write';
import { writeParagraphFormatRevision, writeParagraphMarkRevision } from './write-revisions';
import { paragraphJustification } from './paragraph-alignment';
import { parseDirectParagraphProperties } from './paragraph-properties';
import { parsePropertiesSnapshot } from './revision-properties';

export function setAttribute(element: XmlElement, local: string, value: string): void {
	element.setAttributeNS(WORD_NS, `w:${local}`, value);
}
function removeChildren(element: XmlElement, local: string): void {
	for (const child of children(element, local)) element.removeChild(child);
}

/**
 * Writes `paragraph`'s `w:pPr` into `node` (creating it first when missing), relative to `base`.
 * Returns the element and the properties it was written against: the restored or kept source
 * properties when they replaced the old element, else `base`.
 */
export function writeParagraphMark(
	doc: XmlDocument,
	paragraph: Paragraph,
	node: XmlElement,
	base: Paragraph | undefined,
): { pPr: XmlElement; propertyBase: Paragraph | undefined } {
	let pPr = first(node, 'pPr');
	let propertyBase = base;
	const restoring =
		paragraph.restoredParagraphPropertiesXml &&
		(!base || paragraph.restoredParagraphPropertiesXml !== base.restoredParagraphPropertiesXml);
	const propertiesXml =
		paragraph.restoredParagraphPropertiesXml ?? paragraph.sourceParagraphPropertiesXml;
	if (propertiesXml && (restoring || paragraph.sourceParagraphPropertiesXml || !base)) {
		const restored = parsePropertiesSnapshot(propertiesXml, 'pPr');
		if (pPr) node.removeChild(pPr);
		pPr = importElement(doc, restored);
		node.insertBefore(pPr, node.firstChild);
		propertyBase = {
			type: 'paragraph',
			id: paragraph.id,
			runs: paragraph.runs,
			...parseDirectParagraphProperties(restored),
		};
	}
	if (!pPr) {
		pPr = makeW(doc, 'pPr');
		node.insertBefore(pPr, node.firstChild);
	}
	if (
		!propertyBase ||
		paragraph.align !== propertyBase.align ||
		paragraph.justification !== propertyBase.justification ||
		paragraph.direction !== propertyBase.direction
	) {
		removeChildren(pPr, 'jc');
		const jc = paragraphJustification(paragraph);
		if (jc) {
			const align = makeW(doc, 'jc');
			setAttribute(align, 'val', jc);
			pPr.appendChild(align);
		}
	}
	if (!propertyBase || paragraph.style !== propertyBase.style) {
		removeChildren(pPr, 'pStyle');
		if (paragraph.style) {
			const style = makeW(doc, 'pStyle');
			setAttribute(style, 'val', paragraph.style);
			pPr.appendChild(style);
		}
	}
	writeParagraphProperties(doc, pPr, paragraph, propertyBase);
	writeNumberingProperties(doc, pPr, paragraph, propertyBase);
	writeParagraphMarkRevision(doc, pPr, paragraph, restoring ? undefined : base);
	writeParagraphFormatRevision(doc, pPr, paragraph.formatRevision);
	writeTabStops(doc, pPr, paragraph, propertyBase);
	orderParagraphProperties(pPr);
	return { pPr, propertyBase };
}

/**
 * The `w:pPr` a new paragraph is written with, as written into a fresh `w:p` (empty when the
 * paragraph has no properties): what the paragraph writer produces for its mark.
 */
export function newParagraphProperties(doc: XmlDocument, paragraph: Paragraph): XmlElement {
	return writeParagraphMark(doc, paragraph, makeW(doc, 'p'), undefined).pPr;
}
