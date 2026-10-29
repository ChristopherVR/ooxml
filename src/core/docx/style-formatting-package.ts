// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Changes to the paragraph formatting of styles already in `word/styles.xml`: alignment,
// direction, spacing, indents, pagination toggles, borders and shading. The style's other
// children, and every `w:pPr` child the model does not hold, stay as they were; the properties
// are written by the same code as a paragraph's direct formatting.
import type JSZip from 'jszip';
import type { DocumentModel, Paragraph, ParagraphFormatting } from './model';
import { paragraphJustification } from './paragraph-alignment';
import { orderParagraphProperties } from './tab-stops';
import { writeParagraphProperties } from './write-paragraph-properties';
import { buildXml, children, first, getW, makeW, parseXml, WORD_NS, type XmlElement } from './xml';

const STYLES_PART = 'word/styles.xml';
/** Children of `w:style` that follow `w:pPr` in the schema sequence. */
const AFTER_PPR = ['rPr', 'tblPr', 'trPr', 'tcPr', 'tblStylePr'];

/** The formatting keys a style edit can write. Others (the outline level) are still refused. */
export const WRITABLE_STYLE_FORMATTING: readonly (keyof ParagraphFormatting)[] = [
	'align',
	'justification',
	'direction',
	'spacingBeforeTwips',
	'spacingAfterTwips',
	'lineSpacingTwips',
	'lineSpacingRule',
	'indentLeftTwips',
	'indentRightTwips',
	'indentStartTwips',
	'indentEndTwips',
	'firstLineTwips',
	'hangingTwips',
	'pageBreakBefore',
	'keepNext',
	'keepLines',
	'widowControl',
	'contextualSpacing',
	'suppressLineNumbers',
	'borders',
	'shadingFill',
];

/** `formatting` without the keys a style edit can write. */
export function unwritableFormatting(formatting: ParagraphFormatting): ParagraphFormatting {
	const rest: Record<string, unknown> = { ...formatting };
	for (const key of WRITABLE_STYLE_FORMATTING) delete rest[key];
	return rest as ParagraphFormatting;
}

const asParagraph = (formatting: ParagraphFormatting): Paragraph => ({
	type: 'paragraph',
	id: 'style',
	runs: [],
	...formatting,
});

function pPrOf(style: XmlElement): XmlElement {
	let pPr = first(style, 'pPr');
	if (pPr) return pPr;
	pPr = makeW(style.ownerDocument, 'pPr');
	const anchor =
		Array.from(style.childNodes).find(
			(node): node is XmlElement =>
				node.nodeType === 1 && AFTER_PPR.includes((node as XmlElement).localName),
		) ?? null;
	style.insertBefore(pPr, anchor);
	return pPr;
}

/** Writes the changed formatting of existing paragraph styles into `word/styles.xml`. */
export async function applyStyleFormatting(
	zip: JSZip,
	model: DocumentModel,
	base: DocumentModel,
): Promise<void> {
	const next = model.paragraphStyles?.styles ?? {};
	const prior = base.paragraphStyles?.styles ?? {};
	const changed = Object.keys(next).filter(
		(id) => JSON.stringify(next[id]?.formatting) !== JSON.stringify(prior[id]?.formatting),
	);
	if (!changed.length) return;
	const file = zip.file(STYLES_PART);
	if (!file) throw new Error('Cannot edit a style: the package has no styles.xml.');
	const doc = parseXml(await file.async('string'));
	const styles = children(doc.documentElement, 'style');
	for (const id of changed) {
		const style = styles.find(
			(element) => getW(element, 'type') === 'paragraph' && getW(element, 'styleId') === id,
		);
		if (!style) throw new Error(`Cannot edit style "${id}": it is not in styles.xml.`);
		const after = asParagraph(next[id]!.formatting);
		const before = asParagraph(prior[id]?.formatting ?? {});
		const pPr = pPrOf(style);
		if (
			after.align !== before.align ||
			after.justification !== before.justification ||
			after.direction !== before.direction
		) {
			for (const jc of children(pPr, 'jc')) pPr.removeChild(jc);
			const value = paragraphJustification(after);
			if (value) {
				const jc = makeW(doc, 'jc');
				jc.setAttributeNS(WORD_NS, 'w:val', value);
				pPr.appendChild(jc);
			}
		}
		writeParagraphProperties(doc, pPr, after, before);
		orderParagraphProperties(pPr);
	}
	zip.file(STYLES_PART, buildXml(doc));
}
