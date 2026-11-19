/**
 * XML canonicalization and DOM navigation utilities for digital signatures.
 *
 * Node-only - depends on `@xmldom/xmldom` and `xml-crypto`.
 *
 * All public functions use minimal structural (duck-typed) interfaces instead
 * of the standard lib.dom.d.ts types. @xmldom/xmldom's Document/Element/Node
 * types do not extend the standard DOM interfaces in TypeScript 6 (the old DOM
 * lib baked into @xmldom/xmldom's .d.ts predates properties like
 * activeViewTransition), so using the standard types causes TS2345 errors in
 * the DTS build. These structural interfaces are satisfied by both xmldom and
 * standard DOM objects at runtime.
 */

import { DOMParser } from '@xmldom/xmldom';
import { SignedXml } from 'xml-crypto';

import { withoutByteOrderMark } from '../../xml/index';

const ELEMENT_NODE = 1;
const DOCUMENT_TYPE_NODE = 10;

/** Canonical XML 1.0, the default for same-document references (XML-DSig section 4.3.3.2). */
export const INCLUSIVE_C14N = 'http://www.w3.org/TR/2001/REC-xml-c14n-20010315';

/** A DOM element-like collection returned by getElementsByTagName. */
export interface XmlElementCollection {
	readonly length: number;
	item(index: number): XmlElement | null;
	[index: number]: XmlElement;
	[Symbol.iterator](): Iterator<XmlElement>;
}

interface XmlAttribute {
	readonly name: string;
	readonly value: string;
}

interface XmlChildNode {
	readonly nodeType: number;
}

/** Structural interface covering the @xmldom/xmldom Element operations used here. */
export interface XmlElement extends XmlChildNode {
	readonly nodeName: string;
	readonly localName?: string | null;
	readonly namespaceURI?: string | null;
	readonly prefix?: string | null;
	readonly textContent: string | null;
	readonly parentNode: XmlElement | null;
	readonly attributes: ArrayLike<XmlAttribute>;
	readonly childNodes: ArrayLike<XmlChildNode>;
	getAttribute(name: string): string | null;
	getElementsByTagName(localName: string): XmlElementCollection;
	getElementsByTagNameNS(namespaceURI: string | null, localName: string): XmlElementCollection;
}

/** Structural interface covering @xmldom/xmldom Document operations used here. */
export interface XmlDocument {
	readonly documentElement: XmlElement | null;
	readonly childNodes: ArrayLike<XmlChildNode>;
	getElementsByTagName(localName: string): XmlElementCollection;
	getElementsByTagNameNS(namespaceURI: string | null, localName: string): XmlElementCollection;
}

/**
 * Parse XML that decides a security verdict. Malformed markup and document type declarations are
 * errors: entity definitions could make the bytes a signer saw differ from the tree read here. A
 * leading byte order mark is encoding information, so it is ignored.
 */
export function parseStrictXml(xml: string): XmlDocument {
	const parser = new DOMParser({
		onError: (level: string, message: string) => {
			if (level !== 'warning') {
				throw new Error(`Invalid XML: ${message}`);
			}
		},
	});
	const doc = parser.parseFromString(
		withoutByteOrderMark(xml),
		'text/xml',
	) as unknown as XmlDocument;
	if (Array.from(doc.childNodes).some((node) => node.nodeType === DOCUMENT_TYPE_NODE)) {
		throw new Error('Invalid XML: document type declarations are not allowed.');
	}
	if (!doc.documentElement) {
		throw new Error('Invalid XML: no document element.');
	}
	return doc;
}

/** The element children of `parent`, in document order. */
export function childElements(parent: XmlDocument | XmlElement): XmlElement[] {
	return Array.from(parent.childNodes).filter(
		(node): node is XmlElement => node.nodeType === ELEMENT_NODE,
	);
}

/** Get the local name of a DOM node, stripping any namespace prefix. */
export function getNodeLocalName(node: XmlElement): string {
	if (node.localName) {
		return node.localName;
	}
	const nodeName = node.nodeName || '';
	const sep = nodeName.indexOf(':');
	return sep >= 0 ? nodeName.slice(sep + 1) : nodeName;
}

/**
 * Find the first descendant element matching a local name,
 * ignoring namespace prefixes.
 */
export function getFirstDescendantElementByLocalName(
	parent: XmlDocument | XmlElement,
	localName: string,
): XmlElement | undefined {
	const elements = parent.getElementsByTagName('*');
	for (let index = 0; index < elements.length; index += 1) {
		const element = elements.item(index);
		if (!element) {
			continue;
		}
		if (getNodeLocalName(element) === localName) {
			return element;
		}
	}
	return undefined;
}

const namespaceDeclarationPrefix = (attributeName: string): string | undefined =>
	attributeName === 'xmlns'
		? ''
		: attributeName.startsWith('xmlns:')
			? attributeName.slice('xmlns:'.length)
			: undefined;

/**
 * The namespace declarations `element` inherits from its ancestors. Canonical XML renders them on
 * the apex of a document subset, and canonicalizing a detached copy would lose them.
 */
function inheritedNamespaces(element: XmlElement): { prefix: string; namespaceURI: string }[] {
	const seen = new Set<string>([element.prefix ?? '']);
	for (const attribute of Array.from(element.attributes ?? [])) {
		const prefix = namespaceDeclarationPrefix(attribute.name);
		if (prefix !== undefined) seen.add(prefix);
	}
	const inherited: { prefix: string; namespaceURI: string }[] = [];
	let ancestor = element.parentNode;
	while (ancestor && ancestor.nodeType === ELEMENT_NODE) {
		for (const attribute of Array.from(ancestor.attributes)) {
			const prefix = namespaceDeclarationPrefix(attribute.name);
			if (prefix === undefined || seen.has(prefix)) continue;
			seen.add(prefix);
			if (attribute.value) inherited.push({ prefix, namespaceURI: attribute.value });
		}
		ancestor = ancestor.parentNode;
	}
	return inherited;
}

/**
 * Canonicalize a DOM node using the specified canonicalization algorithm, in the context of its
 * document. Delegates to xml-crypto's C14N implementation.
 */
export function canonicalizeNode(node: XmlDocument | XmlElement, algorithm: string): string {
	const canonicalizer = new SignedXml();
	const ancestorNamespaces = 'parentNode' in node ? inheritedNamespaces(node) : [];
	// xmldom nodes are not assignable to lib.dom.d.ts Node in TypeScript 6.
	// eslint-disable-next-line @typescript-eslint/no-explicit-any
	return canonicalizer.getCanonXml([algorithm], node as any, { ancestorNamespaces });
}

/**
 * Canonicalize a `<SignedInfo>` XML fragment for signing.
 * Uses Exclusive XML Canonicalization (exc-c14n#).
 */
export function canonicalizeSignedInfoXml(signedInfoXml: string): string {
	return canonicalizeNode(
		parseStrictXml(signedInfoXml).documentElement!,
		'http://www.w3.org/2001/10/xml-exc-c14n#',
	);
}
