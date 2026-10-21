// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Writes section edits: page size/orientation, margins, columns and break type for changed
// sections, new section breaks (a paragraph-level w:sectPr copied from the following section,
// as Word does) and removed breaks. Other section properties stay protected.
import type { Block, HeaderFooterSlots, SectionProperties } from './model';
import { setW, writeSectionProperties } from './write-section-properties';
import { orderSectionProperties } from './element-order';
import { orderParagraphProperties } from './tab-stops';
import { flatBlocks } from './parse-content-control';
import { children, first, makeW, WORD_NS, type XmlDocument, type XmlElement } from './xml';

/** Section properties this writer can change; any other difference is rejected. */
const WRITABLE = new Set([
	'endsAtBlockId',
	'type',
	'pageWidthTwips',
	'pageHeightTwips',
	'orientation',
	'marginTopTwips',
	'marginRightTwips',
	'marginBottomTwips',
	'marginLeftTwips',
	'headerDistanceTwips',
	'footerDistanceTwips',
	'gutterTwips',
	'columns',
	'footnoteColumns',
	'titlePage',
	'verticalAlign',
	'pageNumbering',
	'lineNumbering',
	'lineNumberSettings',
	'pageBorders',
	'headers',
	'footers',
]);

function assertWritable(section: SectionProperties, base: SectionProperties | undefined) {
	if (!base) return;
	for (const key of new Set([...Object.keys(section), ...Object.keys(base)]))
		if (
			!WRITABLE.has(key) &&
			JSON.stringify(section[key as keyof SectionProperties]) !==
				JSON.stringify(base[key as keyof SectionProperties])
		)
			throw new Error(
				`Editing the section property "${key}" is not supported; only page size, orientation, margins, columns, break type, vertical alignment, first-page setting and page numbering can change.`,
			);
}

const REL_NAMESPACE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/**
 * Adds `w:headerReference` / `w:footerReference` for header and footer parts this section gained
 * (`references` maps each new part to its relationship id), removing references when a story links
 * to the previous section. Unchanged references stay as they are.
 */
function writeNewReferences(
	doc: XmlDocument,
	sectPr: XmlElement,
	section: SectionProperties,
	previous: SectionProperties | undefined,
	references: ReadonlyMap<string, string>,
): void {
	for (const [kind, name] of [
		['headers', 'headerReference'],
		['footers', 'footerReference'],
	] as const) {
		// Remove deleted local stories and references copied from following geometry into new breaks.
		if (!previous || previous[kind])
			for (const reference of children(sectPr, name)) {
				const slot = reference.getAttributeNS(WORD_NS, 'type') as keyof HeaderFooterSlots;
				if (!section[kind]?.[slot]) sectPr.removeChild(reference);
			}
		for (const [slot, content] of Object.entries(section[kind] ?? {}) as Array<
			[keyof HeaderFooterSlots, HeaderFooterSlots[keyof HeaderFooterSlots]]
		>) {
			const id = content?.partName ? references.get(content.partName) : undefined;
			if (!id || previous?.[kind]?.[slot]?.partName === content?.partName) continue;
			for (const old of children(sectPr, name))
				if (old.getAttributeNS(WORD_NS, 'type') === slot) sectPr.removeChild(old);
			const reference = makeW(doc, name);
			setW(reference, 'type', slot);
			reference.setAttributeNS(REL_NAMESPACE, 'r:id', id);
			sectPr.appendChild(reference);
		}
	}
}

/** The paragraph-level w:sectPr of a top-level paragraph, if any. */
const paragraphSection = (node: XmlElement | undefined) =>
	node?.localName === 'p' ? first(first(node, 'pPr'), 'sectPr') : undefined;

/**
 * Applies `sections` to a body already rewritten for `blocks`. `base` is the loaded section list
 * (empty for new documents); unchanged sections are left byte-identical.
 */
export function applySectionEdits(
	doc: XmlDocument,
	body: XmlElement,
	blocks: Block[],
	sections: SectionProperties[],
	base: SectionProperties[],
	/** Relationship ids of header/footer parts created for this save, by part name. */
	references: ReadonlyMap<string, string> = new Map(),
): void {
	if (!sections.length) {
		if (base.length)
			throw new Error('A document needs at least one section; the section list cannot be emptied.');
		return;
	}
	// The body's paragraphs and tables as the model's blocks are read, through content controls.
	const topLevel = flatBlocks(body).map(({ element }) => element);
	const nodeFor = (id: string) => topLevel[blocks.findIndex((block) => block.id === id)];
	const ends = new Set(sections.slice(0, -1).map((section) => section.endsAtBlockId));
	// Remove breaks the model no longer has.
	blocks.forEach((block, index) => {
		const sectPr = paragraphSection(topLevel[index]);
		if (sectPr && !ends.has(block.id)) sectPr.parentNode?.removeChild(sectPr);
	});
	let bodySection = children(body, 'sectPr').at(-1);
	if (!bodySection) {
		bodySection = makeW(doc, 'sectPr');
		body.appendChild(bodySection);
	}
	const baseByEnd = new Map(base.map((section) => [section.endsAtBlockId, section]));
	// Walk backwards so each new break can copy the section that follows it.
	let following = bodySection;
	for (let index = sections.length - 1; index >= 0; index--) {
		const section = sections[index];
		if (!section) continue;
		const last = index === sections.length - 1;
		let sectPr: XmlElement | undefined = last
			? bodySection
			: paragraphSection(nodeFor(section.endsAtBlockId));
		const previous = last ? base.at(-1) : baseByEnd.get(section.endsAtBlockId);
		if (!sectPr) {
			const paragraph = nodeFor(section.endsAtBlockId);
			if (paragraph?.localName !== 'p')
				throw new Error('A section break must end on a paragraph, not a table.');
			let pPr = first(paragraph, 'pPr');
			if (!pPr) {
				pPr = makeW(doc, 'pPr');
				paragraph.insertBefore(pPr, paragraph.firstChild);
			}
			sectPr = following.cloneNode(true) as XmlElement;
			pPr.appendChild(sectPr);
			orderParagraphProperties(pPr);
		}
		assertWritable(section, previous);
		if (!previous || JSON.stringify(previous) !== JSON.stringify(section)) {
			writeSectionProperties(doc, sectPr, section, previous);
			writeNewReferences(doc, sectPr, section, previous, references);
			orderSectionProperties(sectPr);
		}
		following = sectPr;
	}
}
