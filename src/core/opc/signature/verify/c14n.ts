/**
 * Canonical XML 1.0 (W3C REC-xml-c14n-20010315) and Exclusive XML Canonicalization 1.0 (W3C
 * REC-xml-exc-c14n-20020718), with and without comments, over a parsed DOM. Pure TypeScript, so a
 * signature verifies the same way in a browser and in Node.
 *
 * The input is a whole document or the subtree of one element in its document (the node-set of a
 * same-document `#id` reference). Namespace declarations and, for Canonical XML 1.0, `xml:`
 * attributes that the subtree inherits from ancestors outside it are rendered on its apex.
 */

import {
	EXCLUSIVE_C14N,
	EXCLUSIVE_C14N_WITH_COMMENTS,
	INCLUSIVE_C14N,
	INCLUSIVE_C14N_WITH_COMMENTS,
} from '../constants';

const ELEMENT_NODE = 1;
const TEXT_NODE = 3;
const CDATA_SECTION_NODE = 4;
const PROCESSING_INSTRUCTION_NODE = 7;
const COMMENT_NODE = 8;
const DOCUMENT_NODE = 9;

const XMLNS_NS = 'http://www.w3.org/2000/xmlns/';
const XML_NS = 'http://www.w3.org/XML/1998/namespace';

interface Mode {
	exclusive: boolean;
	comments: boolean;
}

const MODES: ReadonlyMap<string, Mode> = new Map([
	[INCLUSIVE_C14N, { exclusive: false, comments: false }],
	[INCLUSIVE_C14N_WITH_COMMENTS, { exclusive: false, comments: true }],
	[EXCLUSIVE_C14N, { exclusive: true, comments: false }],
	[EXCLUSIVE_C14N_WITH_COMMENTS, { exclusive: true, comments: true }],
]);

/** Whether {@link canonicalize} implements `algorithm`. */
export const isCanonicalizationAlgorithm = (algorithm: string): boolean => MODES.has(algorithm);

const escapeText = (value: string): string =>
	value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/\r/g, '&#xD;');

const escapeAttribute = (value: string): string =>
	value
		.replace(/&/g, '&amp;')
		.replace(/</g, '&lt;')
		.replace(/"/g, '&quot;')
		.replace(/\t/g, '&#x9;')
		.replace(/\n/g, '&#xA;')
		.replace(/\r/g, '&#xD;');

/** The prefix a namespace declaration attribute declares (`''` for the default), else undefined. */
function declaredPrefix(attribute: Attr): string | undefined {
	if (attribute.name === 'xmlns') return '';
	if (attribute.name.startsWith('xmlns:')) return attribute.name.slice('xmlns:'.length);
	return attribute.namespaceURI === XMLNS_NS ? (attribute.localName ?? '') : undefined;
}

const isElement = (node: Node | null): node is Element => node?.nodeType === ELEMENT_NODE;

/** The namespaces in scope on `element`: its own declarations, then each ancestor's. */
function namespacesInScope(element: Element): Map<string, string> {
	const scope = new Map<string, string>();
	for (let current: Node | null = element; isElement(current); current = current.parentNode) {
		for (const attribute of Array.from(current.attributes)) {
			const prefix = declaredPrefix(attribute);
			if (prefix !== undefined && prefix !== 'xml' && !scope.has(prefix))
				scope.set(prefix, attribute.value);
		}
	}
	return scope;
}

/** `xml:` attributes the apex inherits from ancestors outside the node-set (C14N 1.0, 2.4). */
function inheritedXmlAttributes(apex: Element): Attr[] {
	const own = new Set(
		Array.from(apex.attributes)
			.filter((attribute) => attribute.namespaceURI === XML_NS)
			.map((attribute) => attribute.localName),
	);
	const inherited: Attr[] = [];
	for (let current = apex.parentNode; isElement(current); current = current.parentNode) {
		for (const attribute of Array.from(current.attributes)) {
			if (attribute.namespaceURI !== XML_NS || own.has(attribute.localName)) continue;
			own.add(attribute.localName);
			inherited.push(attribute);
		}
	}
	return inherited;
}

const compareStrings = (left: string, right: string): number =>
	left < right ? -1 : left > right ? 1 : 0;

const prefixOf = (qualifiedName: string): string => {
	const separator = qualifiedName.indexOf(':');
	return separator < 0 ? '' : qualifiedName.slice(0, separator);
};

class Canonicalizer {
	private out = '';

	constructor(private readonly mode: Mode) {}

	run(node: Document | Element): string {
		if (node.nodeType === DOCUMENT_NODE) this.document(node as Document);
		else this.element(node as Element, new Map(), true);
		return this.out;
	}

	private document(document: Document): void {
		let seenRoot = false;
		for (const child of Array.from(document.childNodes)) {
			if (isElement(child)) {
				this.element(child, new Map(), false);
				seenRoot = true;
			} else if (this.renders(child)) {
				if (seenRoot) this.out += '\n';
				this.leaf(child);
				if (!seenRoot) this.out += '\n';
			}
		}
	}

	private renders(node: Node): boolean {
		// The parser reports the XML declaration as an instruction named `xml`; it is not one.
		return (
			(node.nodeType === PROCESSING_INSTRUCTION_NODE &&
				(node as ProcessingInstruction).target !== 'xml') ||
			(node.nodeType === COMMENT_NODE && this.mode.comments)
		);
	}

	private leaf(node: Node): void {
		if (node.nodeType === COMMENT_NODE) {
			this.out += `<!--${node.nodeValue ?? ''}-->`;
			return;
		}
		const instruction = node as ProcessingInstruction;
		this.out += instruction.data
			? `<?${instruction.target} ${instruction.data}?>`
			: `<?${instruction.target}?>`;
	}

	/** The namespace declarations to render on `element`, sorted by prefix (default first). */
	private namespaces(
		element: Element,
		attributes: readonly Attr[],
		rendered: ReadonlyMap<string, string>,
	): [string, string][] {
		const scope = namespacesInScope(element);
		let prefixes: Iterable<string> = scope.keys();
		if (this.mode.exclusive) {
			const used = new Set([prefixOf(element.nodeName)]);
			for (const attribute of attributes) {
				const prefix = prefixOf(attribute.name);
				if (prefix !== '' && prefix !== 'xml') used.add(prefix);
			}
			prefixes = used;
		}
		const result: [string, string][] = [];
		for (const prefix of prefixes) {
			const uri = scope.get(prefix) ?? '';
			// An empty default namespace is only stated to undo a non-empty one rendered above.
			if ((rendered.get(prefix) ?? '') === uri) continue;
			if (prefix !== '' && uri === '') continue;
			result.push([prefix, uri]);
		}
		return result.sort(([left], [right]) => compareStrings(left, right));
	}

	private element(element: Element, rendered: ReadonlyMap<string, string>, apex: boolean): void {
		const attributes = Array.from(element.attributes).filter(
			(attribute) => declaredPrefix(attribute) === undefined,
		);
		if (apex && !this.mode.exclusive) attributes.push(...inheritedXmlAttributes(element));
		const namespaces = this.namespaces(element, attributes, rendered);
		const childRendered = new Map(rendered);
		for (const [prefix, uri] of namespaces) childRendered.set(prefix, uri);

		this.out += `<${element.nodeName}`;
		for (const [prefix, uri] of namespaces)
			this.out += `${prefix === '' ? ' xmlns' : ` xmlns:${prefix}`}="${escapeAttribute(uri)}"`;
		const sorted = attributes
			.map((attribute) => ({
				attribute,
				namespace: attribute.namespaceURI ?? '',
				local: attribute.localName ?? attribute.name,
			}))
			.sort(
				(left, right) =>
					compareStrings(left.namespace, right.namespace) ||
					compareStrings(left.local, right.local),
			);
		for (const { attribute } of sorted)
			this.out += ` ${attribute.name}="${escapeAttribute(attribute.value)}"`;
		this.out += '>';

		for (const child of Array.from(element.childNodes)) {
			if (isElement(child)) this.element(child, childRendered, false);
			else if (child.nodeType === TEXT_NODE || child.nodeType === CDATA_SECTION_NODE)
				this.out += escapeText(child.nodeValue ?? '');
			else if (this.renders(child)) this.leaf(child);
		}
		this.out += `</${element.nodeName}>`;
	}
}

/**
 * Canonicalize `node` with the canonicalization `algorithm` (one of the four URIs above).
 *
 * @throws When the algorithm is not one of them.
 */
export function canonicalize(node: Document | Element, algorithm: string): string {
	const mode = MODES.get(algorithm);
	if (!mode) throw new Error(`Unsupported canonicalization algorithm ${algorithm}.`);
	return new Canonicalizer(mode).run(node);
}
