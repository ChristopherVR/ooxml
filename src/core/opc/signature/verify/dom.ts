/** DOM reading shared by the signature verifier. */

import { parseXml } from '../../../xml/index';

const ELEMENT_NODE = 1;

/**
 * Parse XML that decides a security verdict. Malformed markup and document type declarations are
 * errors: entity definitions could make the bytes a signer saw differ from the tree read here. A
 * leading byte order mark is encoding information, so it is ignored.
 */
export const parseStrictXml = (xml: string): Document =>
	parseXml(xml, { label: 'signature' }) as unknown as Document;

/** The element children of `parent`, in document order. */
export const childElements = (parent: Document | Element): Element[] =>
	Array.from(parent.childNodes).filter((node): node is Element => node.nodeType === ELEMENT_NODE);

/** The local name of `element`, without its prefix. */
export function localNameOf(element: Element): string {
	if (element.localName) return element.localName;
	const separator = element.nodeName.indexOf(':');
	return separator >= 0 ? element.nodeName.slice(separator + 1) : element.nodeName;
}

/** The first descendant element named `localName` in any namespace. */
export const firstDescendantNamed = (
	parent: Document | Element,
	localName: string,
): Element | undefined =>
	Array.from(parent.getElementsByTagName('*')).find(
		(element) => localNameOf(element) === localName,
	);
