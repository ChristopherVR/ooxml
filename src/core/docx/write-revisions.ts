// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import type { Paragraph, Revision } from './model';
import {
	children,
	first,
	makeW,
	type XmlDocument,
	type XmlElement,
	WORD_NS,
	WORD_DATE_UTC_NS,
} from './xml';
import { parsePropertiesSnapshot } from './revision-properties';

const REVISION_WRAPPER_NAMES = ['ins', 'del', 'moveFrom', 'moveTo'];
/** Range markers the writer regenerates from the model: comment anchors and move ranges. */
const COMMENT_ANCHOR_NAMES = [
	'commentRangeStart',
	'commentRangeEnd',
	'commentReference',
	'moveFromRangeStart',
	'moveFromRangeEnd',
	'moveToRangeStart',
	'moveToRangeEnd',
];

function isWordElement(element: XmlElement, names: string[]): boolean {
	return (
		(!element.namespaceURI || element.namespaceURI === WORD_NS) && names.includes(element.localName)
	);
}
export const isRevisionWrapperElement = (element: XmlElement): boolean =>
	isWordElement(element, REVISION_WRAPPER_NAMES);
export const isCommentAnchorElement = (element: XmlElement): boolean =>
	isWordElement(element, COMMENT_ANCHOR_NAMES);

function setAttribute(element: XmlElement, local: string, value: string): void {
	element.setAttributeNS(WORD_NS, `w:${local}`, value);
}
function removeChildren(element: XmlElement, local: string): void {
	for (const child of children(element, local)) element.removeChild(child);
}

function writeRevisionMetadata(element: XmlElement, revision: Revision): void {
	setAttribute(element, 'id', revision.id);
	setAttribute(element, 'author', revision.author);
	if (revision.date) setAttribute(element, 'date', revision.date);
	else element.removeAttributeNS(WORD_NS, 'date');
	if (revision.dateUtc) element.setAttributeNS(WORD_DATE_UTC_NS, 'w16du:dateUtc', revision.dateUtc);
	else element.removeAttributeNS(WORD_DATE_UTC_NS, 'dateUtc');
}
function revisionTag(kind: Revision['kind']): string {
	return kind === 'delete'
		? 'del'
		: kind === 'moveFrom'
			? 'moveFrom'
			: kind === 'moveTo'
				? 'moveTo'
				: 'ins';
}
/** Rebuild historical run formatting when a text edit requires a fresh run. */
export function writeRunFormatRevision(
	doc: XmlDocument,
	props: XmlElement,
	revision: Revision | undefined,
): void {
	writeFormatRevision(doc, props, revision?.kind === 'formatChange' ? revision : undefined, 'rPr');
}

/** Rebuild prior paragraph properties when a paragraph is edited. */
export function writeParagraphFormatRevision(
	doc: XmlDocument,
	props: XmlElement,
	revision: Revision | undefined,
): void {
	writeFormatRevision(
		doc,
		props,
		revision?.kind === 'paragraphChange' ? revision : undefined,
		'pPr',
	);
}

function writeFormatRevision(
	doc: XmlDocument,
	props: XmlElement,
	revision: Revision | undefined,
	local: 'rPr' | 'pPr',
): void {
	const original = first(props, `${local}Change`);
	removeChildren(props, `${local}Change`);
	if (!revision) return;
	const xml =
		local === 'rPr' ? revision.previousRunPropertiesXml : revision.previousParagraphPropertiesXml;
	if (!xml)
		throw new Error(`Cannot write a formatting revision without its prior ${local} snapshot.`);
	const previous = parsePropertiesSnapshot(xml, local);
	const change = original ?? makeW(doc, `${local}Change`);
	removeChildren(change, local);
	writeRevisionMetadata(change, revision);
	change.appendChild(doc.importNode(previous, true));
	props.appendChild(change);
}
export function revisionWrapper(
	doc: XmlDocument,
	revision: Revision,
	content: XmlElement,
): XmlElement {
	const wrapper = makeW(doc, revisionTag(revision.kind));
	writeRevisionMetadata(wrapper, revision);
	wrapper.appendChild(content);
	return wrapper;
}
/** Deleted text uses `w:delText`; moved-from content retains `w:t`. */
export function convertToDeleteText(doc: XmlDocument, run: XmlElement): void {
	for (const t of children(run, 't')) {
		const delText = doc.createElementNS(WORD_NS, 'w:delText');
		const space = t.getAttributeNS('http://www.w3.org/XML/1998/namespace', 'space');
		if (space) delText.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', space);
		while (t.firstChild) delText.appendChild(t.firstChild);
		run.replaceChild(delText, t);
	}
}

/** Writes/clears the paragraph mark's own tracked insertion/deletion (`w:pPr/w:rPr/w:ins|w:del`). */
export function writeParagraphMarkRevision(
	doc: XmlDocument,
	pPr: XmlElement,
	paragraph: Paragraph,
	base?: Paragraph,
): void {
	if (base && JSON.stringify(paragraph.markRevision) === JSON.stringify(base.markRevision)) return;
	let rPr = first(pPr, 'rPr');
	if (rPr) {
		removeChildren(rPr, 'ins');
		removeChildren(rPr, 'del');
	}
	if (!paragraph.markRevision) {
		if (rPr && !rPr.attributes.length && !rPr.childNodes.length) pPr.removeChild(rPr);
		return;
	}
	if (!rPr) {
		rPr = makeW(doc, 'rPr');
		pPr.appendChild(rPr);
	}
	const revision = paragraph.markRevision;
	const element = makeW(doc, revision.kind === 'delete' ? 'del' : 'ins');
	writeRevisionMetadata(element, revision);
	rPr.insertBefore(element, rPr.firstChild);
}
