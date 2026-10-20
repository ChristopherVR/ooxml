// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Word-specific view of the shared XML model in `../xml/index.js`: the same parser and
// DOM, with the WordprocessingML namespace as the default for the element/attribute helpers.
import {
	NS,
	attr,
	buildXml,
	children as nsChildren,
	elements,
	first as nsFirst,
	isElement,
	makeNS,
	named as nsNamed,
	parseXml as parseOoxml,
	relAttr,
	textContent,
	type XmlDocument,
	type XmlElement,
	type XmlNode,
} from '../xml/index';

export const WORD_NS = NS.w;
/** Word 2023 revision timestamps, separate from legacy `w:date`. */
export const WORD_DATE_UTC_NS = 'http://schemas.microsoft.com/office/word/2023/wordml/word16du';
/** OOXML relationships namespace used for `r:id`/`r:embed` attributes on drawings and hyperlinks. */
export const REL_NS = NS.r;
export type { XmlDocument, XmlElement, XmlNode };

export const parseXml = (xml: string): XmlDocument => parseOoxml(xml, { label: 'DOCX' });
export { buildXml, elements, isElement, makeNS, textContent };

export const named = (element: XmlElement | null | undefined, local: string): boolean =>
	nsNamed(element, local, WORD_NS);
export const children = (parent: ParentNode, local: string): XmlElement[] =>
	nsChildren(parent, local, WORD_NS);
export const first = (
	parent: ParentNode | null | undefined,
	local: string,
): XmlElement | undefined => nsFirst(parent, local, WORD_NS);
export const getW = (element: XmlElement | null | undefined, local: string): string | undefined =>
	attr(element, local, WORD_NS, 'w');
export const getR = (element: XmlElement | null | undefined, local: string): string | undefined =>
	relAttr(element, local);
export const makeW = (doc: XmlDocument, local: string): XmlElement =>
	doc.createElementNS(WORD_NS, `w:${local}`);

const XMLNS = 'http://www.w3.org/2000/xmlns/';

/**
 * Imports `element` and its subtree into `doc`, without the namespace declarations `doc`'s root
 * element already makes. An element serialized on its own, such as a properties snapshot, declares
 * every prefix it uses; inside the document those declarations would only repeat the root's.
 */
export function importElement(doc: XmlDocument, element: XmlElement): XmlElement {
	const imported = doc.importNode(element, true) as XmlElement;
	const root = doc.documentElement;
	const drop = (node: XmlElement) => {
		for (const attribute of Array.from(node.attributes))
			if (
				attribute.namespaceURI === XMLNS &&
				root?.lookupNamespaceURI(attribute.prefix ? attribute.localName : null) === attribute.value
			)
				node.removeAttributeNode(attribute);
		for (const child of Array.from(node.childNodes)) if (isElement(child)) drop(child);
	};
	drop(imported);
	return imported;
}
