// A small, fast reader for large, plain XML fragments (a worksheet's `<sheetData>`): it builds
// lightweight elements that answer the read-only subset of the DOM the readers use
// (`localName`, `namespaceURI`, `childNodes`, `textContent`, `getAttribute`, `hasAttribute`), so
// the same reading code runs over them. It decodes text and attributes the way the DOM parser
// does (line endings, attribute whitespace, the predefined and numeric entity references).
// Element prefixes resolve through the namespaces in scope where the fragment sits. Comments are
// skipped and CDATA sections kept as text, as the DOM's `textContent` reads them. Anything it does
// not handle (namespace declarations inside the fragment, an undeclared prefix, processing
// instructions, a comment XML forbids, unknown entities, malformed markup) makes it return
// `undefined` (the streaming form throws a `LiteDeclinedError`), and the caller parses the XML
// with the full parser instead.
import type { XmlElement } from './xml';

/** A text node is kept as its decoded string; elements are {@link LiteElement}s. */
export type LiteNode = LiteElement | string;

export class LiteElement {
	readonly nodeType = 1;
	readonly childNodes: LiteNode[] = [];
	/** Attribute names and values, alternating (elements carry few attributes). */
	readonly attrs: string[] = [];

	constructor(
		readonly tagName: string,
		readonly localName: string,
		readonly namespaceURI: string | null,
	) {}

	get nodeName(): string {
		return this.tagName;
	}

	get textContent(): string {
		const nodes = this.childNodes;
		if (nodes.length === 1 && typeof nodes[0] === 'string') return nodes[0];
		let out = '';
		for (const node of nodes) out += typeof node === 'string' ? node : node.textContent;
		return out;
	}

	getAttribute(name: string): string | null {
		const attrs = this.attrs;
		for (let i = 0; i < attrs.length; i += 2) if (attrs[i] === name) return attrs[i + 1] ?? null;
		return null;
	}

	hasAttribute(name: string): boolean {
		return this.getAttribute(name) !== null;
	}

	/** Only un-namespaced attributes are known by namespace (no declarations are read). */
	getAttributeNS(ns: string | null, local: string): string | null {
		return ns ? null : this.getAttribute(local);
	}

	/** This element typed as a DOM element for readers written against the DOM subset above. */
	asElement(): XmlElement {
		return this as unknown as XmlElement;
	}
}

const ENTITIES: Record<string, string> = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" };
const REFERENCE = /&(?:#x([0-9A-Fa-f]+)|#([0-9]+)|([A-Za-z]+));/g;

/** Decodes entity references; undefined when an `&` is not a known, well-formed reference. */
function decode(text: string): string | undefined {
	if (!text.includes('&')) return text;
	let ok = true;
	let found = 0;
	const out = text.replace(REFERENCE, (all, hex: string, dec: string, name: string) => {
		found++;
		if (name !== undefined) {
			const value = ENTITIES[name];
			if (value === undefined) ok = false;
			return value ?? all;
		}
		const code = hex !== undefined ? Number.parseInt(hex, 16) : Number(dec);
		if (!(code > 0 && code <= 0x10ffff)) {
			ok = false;
			return all;
		}
		return String.fromCodePoint(code);
	});
	if (!ok) return undefined;
	let ampersands = 0;
	for (let i = text.indexOf('&'); i >= 0; i = text.indexOf('&', i + 1)) ampersands++;
	return ampersands === found ? out : undefined;
}

const isSpace = (code: number): boolean => code === 32 || code === 9 || code === 10 || code === 13;
const hasLineEnds = (xml: string): boolean =>
	xml.includes('\r') || xml.includes('\u0085') || xml.includes('\u2028') || xml.includes('\u2029');

/** Namespaces in scope by prefix (`''` is the default namespace). */
export type LiteNamespaces = Readonly<Record<string, string>>;

/**
 * Thrown while iterating {@link liteChildElements} at markup this reader does not handle, where
 * {@link parseLiteFragment} returns undefined: the caller reads the XML with the full parser.
 */
export class LiteDeclinedError extends Error {
	constructor() {
		super('The XML fragment needs the full parser');
		this.name = 'LiteDeclinedError';
	}
}

/**
 * Parses the content of one element (its children) into the container {@link LiteElement}
 * `container`, resolving element prefixes through `namespaces`. Returns undefined when the
 * fragment uses markup this reader does not handle or is malformed.
 */
export function parseLiteFragment(
	xml: string,
	container: LiteElement,
	namespaces: LiteNamespaces,
): LiteElement | undefined {
	const nodes = topLevelNodes(xml, namespaces);
	for (let next = nodes.next(); ; next = nodes.next()) {
		if (next.done) return next.value ? container : undefined;
		container.childNodes.push(next.value);
	}
}

/**
 * The child elements of a fragment, read as {@link parseLiteFragment} reads them but built one at
 * a time: each is complete when it is yielded and the next is read only when asked for, so a
 * large fragment (the rows of a worksheet) never exists as one tree. Where `parseLiteFragment`
 * returns undefined, the iteration throws a {@link LiteDeclinedError}, possibly after some
 * elements have been yielded.
 */
export function* liteChildElements(
	xml: string,
	namespaces: LiteNamespaces,
): Generator<XmlElement, void, undefined> {
	const nodes = topLevelNodes(xml, namespaces);
	for (let next = nodes.next(); ; next = nodes.next()) {
		if (next.done) {
			if (next.value) return;
			throw new LiteDeclinedError();
		}
		if (typeof next.value !== 'string') yield next.value.asElement();
	}
}

/**
 * The top-level nodes of a fragment in order, each element once its end tag is read. Returns
 * false when the fragment uses markup this reader does not handle or is malformed.
 */
function* topLevelNodes(
	xml: string,
	namespaces: LiteNamespaces,
): Generator<LiteNode, boolean, undefined> {
	const source = hasLineEnds(xml)
		? xml.replace(/\r[\n\u0085]/g, '\n').replace(/[\r\u0085\u2028\u2029]/g, '\n')
		: xml;
	// Top-level nodes collect in `root` and are handed out whenever no element is open.
	const root = new LiteElement('', '', null);
	const level = root.childNodes;
	const stack: LiteElement[] = [root];
	let top = root;
	const n = source.length;
	let pos = 0;
	while (pos < n) {
		if (level.length && stack.length === 1) {
			yield* level;
			level.length = 0;
		}
		const lt = source.indexOf('<', pos);
		const end = lt < 0 ? n : lt;
		if (end > pos) {
			const text = decode(source.slice(pos, end));
			if (text === undefined) return false;
			top.childNodes.push(text);
		}
		if (lt < 0) break;
		const next = source.charCodeAt(lt + 1);
		if (next === 33) {
			const after = skipCommentOrCdata(source, lt, top);
			if (after < 0) return false;
			pos = after;
			continue;
		}
		if (next === 47) {
			const gt = source.indexOf('>', lt);
			if (gt < 0 || stack.length < 2 || source.slice(lt + 2, gt).trimEnd() !== top.tagName)
				return false;
			stack.pop();
			top = stack[stack.length - 1] as LiteElement;
			pos = gt + 1;
			continue;
		}
		const parsed = startTag(source, lt + 1, namespaces);
		if (!parsed) return false;
		top.childNodes.push(parsed.element);
		if (!parsed.selfClosing) {
			stack.push(parsed.element);
			top = parsed.element;
		}
		pos = parsed.end;
	}
	if (stack.length !== 1) return false;
	yield* level;
	return true;
}

/** `parseXml` refuses a part holding these anywhere, even quoted in a comment or CDATA. */
const DECLARATIONS = /<!DOCTYPE|<!ENTITY/i;

/**
 * Reads the comment or CDATA section at `lt` (a `<!`), adding CDATA text to `parent`, and returns
 * the position after it; -1 for other `<!` markup, an unterminated section, a comment with `--`
 * inside or a `-` before its end, which XML forbids, or a section quoting a declaration (the full
 * parser then decides).
 */
function skipCommentOrCdata(source: string, lt: number, parent: LiteElement): number {
	if (source.startsWith('<!--', lt)) {
		const end = source.indexOf('-->', lt + 4);
		if (end < 0) return -1;
		const body = source.slice(lt + 4, end);
		if (body.includes('--') || body.endsWith('-') || DECLARATIONS.test(body)) return -1;
		return end + 3;
	}
	if (source.startsWith('<![CDATA[', lt)) {
		const end = source.indexOf(']]>', lt + 9);
		if (end < 0) return -1;
		const text = source.slice(lt + 9, end);
		if (DECLARATIONS.test(text)) return -1;
		parent.childNodes.push(text);
		return end + 3;
	}
	return -1;
}

/** Reads a start tag from just after its `<`; undefined for anything but a plain element. */
function startTag(
	source: string,
	from: number,
	namespaces: LiteNamespaces,
): { element: LiteElement; selfClosing: boolean; end: number } | undefined {
	const n = source.length;
	let i = from;
	let colon = -1;
	while (i < n) {
		const code = source.charCodeAt(i);
		if (isSpace(code) || code === 62 || code === 47) break;
		if (code === 58) {
			if (colon >= 0) return undefined;
			colon = i;
		} else if (code === 60 || code === 33 || code === 63) return undefined;
		i++;
	}
	if (i === from || colon === from || colon === i - 1) return undefined;
	const tagName = source.slice(from, i);
	const prefix = colon < 0 ? '' : source.slice(from, colon);
	const ns = namespaces[prefix];
	if (ns === undefined && prefix) return undefined;
	const element = new LiteElement(
		tagName,
		colon < 0 ? tagName : source.slice(colon + 1, i),
		ns || null,
	);
	for (;;) {
		while (i < n && isSpace(source.charCodeAt(i))) i++;
		if (i >= n) return undefined;
		const code = source.charCodeAt(i);
		if (code === 62) return { element, selfClosing: false, end: i + 1 };
		if (code === 47) {
			if (source.charCodeAt(i + 1) !== 62) return undefined;
			return { element, selfClosing: true, end: i + 2 };
		}
		const eq = source.indexOf('=', i);
		if (eq < 0) return undefined;
		const name = source.slice(i, eq).trimEnd();
		if (!name || /[\s<>/"']/.test(name) || name.startsWith('xmlns')) return undefined;
		// The DOM parser rejects an attribute prefix that is not declared (`xml` always is).
		const attrPrefix = name.includes(':') ? name.slice(0, name.indexOf(':')) : '';
		if (attrPrefix && attrPrefix !== 'xml' && !namespaces[attrPrefix]) return undefined;
		let q = eq + 1;
		while (q < n && isSpace(source.charCodeAt(q))) q++;
		const quote = source[q];
		if (quote !== '"' && quote !== "'") return undefined;
		const close = source.indexOf(quote, q + 1);
		if (close < 0) return undefined;
		const raw = source.slice(q + 1, close);
		if (raw.includes('<') || element.hasAttribute(name)) return undefined;
		const value = decode(/[\t\n\r]/.test(raw) ? raw.replace(/[\t\n\r]/g, ' ') : raw);
		if (value === undefined) return undefined;
		element.attrs.push(name, value);
		i = close + 1;
		const after = source.charCodeAt(i);
		if (i < n && !isSpace(after) && after !== 62 && after !== 47) return undefined;
	}
}
