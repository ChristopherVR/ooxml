import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { NS } from './namespaces.js';

export type XmlDocument = Document;
export type XmlElement = Element;
export type XmlNode = Node;
type XmldomNode = Parameters<InstanceType<typeof XMLSerializer>['serializeToString']>[0];

export interface ParseOptions {
	/** Names the document kind in error messages (`Invalid DOCX XML: ...`); defaults to `OOXML`. */
	label?: string;
}

/**
 * Parses an XML part into a DOM document. Parse errors throw (warnings do not), a missing root
 * throws, and DTD or entity declarations are rejected outright (no external or expanded entities).
 */
export function parseXml(xml: string, options: ParseOptions = {}): XmlDocument {
	const label = options.label ?? 'OOXML';
	if (/<!DOCTYPE|<!ENTITY/i.test(xml))
		throw new Error(`${label} XML with DTD or entity declarations is not supported`);
	let parseError = '';
	let parsed: XmlDocument;
	try {
		parsed = new DOMParser({
			onError: (level, message) => {
				if (level !== 'warning') parseError = message;
			},
		}).parseFromString(xml, 'application/xml') as unknown as XmlDocument;
	} catch (error) {
		throw new Error(
			`Invalid ${label} XML: ${error instanceof Error ? error.message : String(error)}`,
		);
	}
	if (parseError || !parsed.documentElement)
		throw new Error(`Invalid ${label} XML: ${parseError || 'missing root element'}`);
	return parsed;
}

export const buildXml = (document: XmlDocument): string =>
	new XMLSerializer().serializeToString(document as unknown as XmldomNode);

export const isElement = (node: Node): node is XmlElement => node.nodeType === 1;
export const elements = (parent: ParentNode): XmlElement[] =>
	Array.from(parent.childNodes).filter(isElement);
export const textContent = (element: XmlElement): string => element.textContent ?? '';
export const makeNS = (doc: XmlDocument, ns: string, qualified: string): XmlElement =>
	doc.createElementNS(ns, qualified);

/**
 * Whether `element` has local name `local` in namespace `ns`. With `lenient`, an element with no
 * namespace also matches (hand-written test XML and some tools omit declarations).
 */
export function named(
	element: XmlElement | null | undefined,
	local: string,
	ns: string,
	lenient = true,
): boolean {
	if (!element || element.localName !== local) return false;
	return element.namespaceURI === ns || (lenient && !element.namespaceURI);
}

/** Direct child elements named `local` in `ns`. */
export const children = (
	parent: ParentNode,
	local: string,
	ns: string,
	lenient = true,
): XmlElement[] => elements(parent).filter((element) => named(element, local, ns, lenient));

/** The first direct child element named `local` in `ns`. */
export const first = (
	parent: ParentNode | null | undefined,
	local: string,
	ns: string,
	lenient = true,
): XmlElement | undefined => (parent ? children(parent, local, ns, lenient)[0] : undefined);

/** An attribute by namespace and local name, falling back to the `prefix:local` form. */
export function attr(
	element: XmlElement | null | undefined,
	local: string,
	ns: string,
	prefix: string,
): string | undefined {
	return (
		element?.getAttributeNS(ns, local) ?? element?.getAttribute(`${prefix}:${local}`) ?? undefined
	);
}

/** `r:id`-style relationship attribute (`id`, `embed`, `link`...) on an element. */
export const relAttr = (element: XmlElement | null | undefined, local: string) =>
	attr(element, local, NS.r, 'r');
