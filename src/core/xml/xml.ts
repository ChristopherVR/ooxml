import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { NS } from './namespaces';

export type XmlDocument = Document;
export type XmlElement = Element;
export type XmlNode = Node;
type XmldomNode = Parameters<InstanceType<typeof XMLSerializer>['serializeToString']>[0];

export interface XmlParseLimits {
	/** Maximum element nesting depth; defaults to {@link DEFAULT_XML_LIMITS}. */
	maxDepth?: number;
	/** Maximum number of element start tags in one part; defaults to {@link DEFAULT_XML_LIMITS}. */
	maxNodes?: number;
}

export interface ParseOptions {
	/** Names the document kind in error messages (`Invalid DOCX XML: ...`); defaults to `OOXML`. */
	label?: string;
	/** Resource limits checked by a linear pre-scan before the DOM is built. */
	limits?: XmlParseLimits;
}

/** Defaults far above any real OOXML part, low enough to stop stack and memory exhaustion. */
export const DEFAULT_XML_LIMITS: Readonly<Required<XmlParseLimits>> = {
	maxDepth: 512,
	maxNodes: 5_000_000,
};

// Tags (with quoted attribute values) give: group 1 `/` for end tags, group 2 the name, group 3
// `/` for empty-element tags. Comments, CDATA and processing instructions are skipped whole with
// `indexOf` (see MARKUP_SKIPS), so their text never counts as markup and no scan is quadratic.
// Names and values never contain `<` (XML forbids it), so one attempt cannot run past the next
// `<` and a run of unterminated tags is scanned once, not once per tag.
const TAG_RE = /<(\/?)([^\s\/><]+)(?:\s+[^\s=\/><]+\s*=\s*(?:"[^"<]*"|'[^'<]*'))*\s*(\/?)>/y;
const MARKUP_SKIPS: readonly (readonly [open: string, close: string])[] = [
	['<!--', '-->'],
	['<![CDATA[', ']]>'],
	['<?', '?>'],
];

/**
 * Linear pre-scan of the markup: throws when nesting depth or the element count exceeds `limits`.
 * It does not check well-formedness; the DOM parser still does that afterwards.
 */
export function checkXmlLimits(xml: string, label: string, limits: XmlParseLimits = {}): void {
	const maxDepth = limits.maxDepth ?? DEFAULT_XML_LIMITS.maxDepth;
	const maxNodes = limits.maxNodes ?? DEFAULT_XML_LIMITS.maxNodes;
	let depth = 0;
	let nodes = 0;
	let at = xml.indexOf('<');
	while (at !== -1) {
		const skip = MARKUP_SKIPS.find(([open]) => xml.startsWith(open, at));
		if (skip) {
			const end = xml.indexOf(skip[1], at + skip[0].length);
			if (end === -1) return; // unterminated: the DOM parser rejects it
			at = xml.indexOf('<', end + skip[1].length);
			continue;
		}
		TAG_RE.lastIndex = at;
		const match = TAG_RE.exec(xml);
		if (!match) {
			at = xml.indexOf('<', at + 1);
			continue;
		}
		const [, closing, , selfClosing] = match;
		at = xml.indexOf('<', at + match[0].length);
		if (closing) {
			depth--;
			continue;
		}
		nodes++;
		if (nodes > maxNodes) throw new Error(`${label} XML exceeds the limit of ${maxNodes} elements`);
		if (!selfClosing && ++depth > maxDepth)
			throw new Error(`${label} XML exceeds the nesting depth limit of ${maxDepth}`);
	}
}

// Parsing turns the CRLF that Office writes after a part's XML declaration into LF (XML
// end-of-line handling); a document parsed from such a part is written back with its CRLF.
const CRLF_DECLARATION = /^<\?xml[^>]*\?>\r\n/;
const crlfDeclarations = new WeakSet<XmlNode>();

/**
 * Parses an XML part into a DOM document. Parse errors throw (warnings do not), a missing root
 * throws, and DTD or entity declarations are rejected outright (no external or expanded entities).
 */
export function parseXml(xml: string, options: ParseOptions = {}): XmlDocument {
	const label = options.label ?? 'OOXML';
	if (/<!DOCTYPE|<!ENTITY/i.test(xml))
		throw new Error(`${label} XML with DTD or entity declarations is not supported`);
	checkXmlLimits(xml, label, options.limits);
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
			{ cause: error },
		);
	}
	if (parseError || !parsed.documentElement)
		throw new Error(`Invalid ${label} XML: ${parseError || 'missing root element'}`);
	if (CRLF_DECLARATION.test(xml)) crlfDeclarations.add(parsed);
	return parsed;
}

export const buildXml = (node: XmlNode): string => {
	const xml = new XMLSerializer().serializeToString(node as unknown as XmldomNode);
	return crlfDeclarations.has(node) ? xml.replace(/^(<\?xml[^>]*\?>)\n/, '$1\r\n') : xml;
};

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
