/**
 * Byte-preserving edits to a flat property part (`docProps/core.xml`, `docProps/app.xml`).
 *
 * The writers decide what changes on the parsed DOM, then splice only those changes into the
 * source text: every byte the edit does not touch (the XML declaration and its line break,
 * indentation, entity and empty-element forms, unknown elements) is kept exactly. Serialising
 * the DOM instead would normalise all of them on every save.
 */
import type { XmlElement } from '../../xml/index';
import { elements } from '../../xml/index';

interface Span {
	start: number;
	/** Just past the start tag's `>`. */
	openEnd: number;
	/** Index of the end tag's `<` (equal to `openEnd` when self-closing). */
	closeStart: number;
	end: number;
	selfClosing: boolean;
}

interface Edit {
	start: number;
	end: number;
	text: string;
	seq: number;
}

const XMLNS = 'http://www.w3.org/2000/xmlns/';

/** Index just past the `>` closing the tag that opens at `from`, honouring quoted values. */
function tagEnd(xml: string, from: number): number {
	let quote = '';
	for (let i = from + 1; i < xml.length; i++) {
		const ch = xml[i];
		if (quote) {
			if (ch === quote) quote = '';
		} else if (ch === '"' || ch === "'") quote = ch;
		else if (ch === '>') return i + 1;
	}
	throw new Error('Unterminated tag in property part');
}

function skipMarkup(xml: string, at: number): number | undefined {
	if (xml.startsWith('<!--', at)) return xml.indexOf('-->', at) + 3;
	if (xml.startsWith('<![CDATA[', at)) return xml.indexOf(']]>', at) + 3;
	if (xml.startsWith('<?', at)) return xml.indexOf('?>', at) + 2;
	return undefined;
}

/** The root element's span and the spans of its element children, in document order. */
function scan(xml: string): { root: Span; children: Span[] } {
	let i = 0;
	let rootStart = -1;
	while (rootStart < 0) {
		const lt = xml.indexOf('<', i);
		if (lt < 0) throw new Error('Property part has no root element');
		const skipped = skipMarkup(xml, lt);
		if (skipped !== undefined) i = skipped;
		else rootStart = lt;
	}
	const openEnd = tagEnd(xml, rootStart);
	if (xml[openEnd - 2] === '/') {
		const root = {
			start: rootStart,
			openEnd,
			closeStart: openEnd,
			end: openEnd,
			selfClosing: true,
		};
		return { root, children: [] };
	}
	const children: Span[] = [];
	let depth = 0;
	let child: { start: number; openEnd: number } | undefined;
	i = openEnd;
	for (;;) {
		const lt = xml.indexOf('<', i);
		if (lt < 0) throw new Error('Property part root is not closed');
		const skipped = skipMarkup(xml, lt);
		if (skipped !== undefined) {
			i = skipped;
			continue;
		}
		const end = tagEnd(xml, lt);
		i = end;
		if (xml[lt + 1] === '/') {
			if (depth === 0) {
				const root = { start: rootStart, openEnd, closeStart: lt, end, selfClosing: false };
				return { root, children };
			}
			depth--;
			if (depth === 0 && child)
				children.push({ ...child, closeStart: lt, end, selfClosing: false });
		} else if (xml[end - 2] === '/') {
			if (depth === 0)
				children.push({ start: lt, openEnd: end, closeStart: end, end, selfClosing: true });
		} else {
			if (depth === 0) child = { start: lt, openEnd: end };
			depth++;
		}
	}
}

/** Element text: `&`, `<`, `>` escaped, and CR kept as a reference so it survives a reload. */
export const escapeText = (text: string): string =>
	text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/\r/g, '&#xD;');

const qualifiedName = (xml: string, span: Span): string =>
	/^<([^\s/>]+)/.exec(xml.slice(span.start, span.openEnd))?.[1] ?? '';

/** A set of edits against one source part, keyed by the parsed root's element children. */
export class PartPatch {
	private readonly root: Span;
	private readonly spans = new Map<XmlElement, Span>();
	private readonly edits: Edit[] = [];
	private readonly appended: string[] = [];
	private readonly prefixes = new Map<string, string>();
	private defaultNs: string | null = null;
	private seq = 0;

	public constructor(
		private readonly xml: string,
		rootElement: XmlElement,
	) {
		const { root, children } = scan(xml);
		const kids = elements(rootElement);
		if (kids.length !== children.length) throw new Error('Property part markup is not supported');
		kids.forEach((kid, index) => this.spans.set(kid, children[index] as Span));
		this.root = root;
		for (const attribute of Array.from(rootElement.attributes)) {
			if (attribute.name === 'xmlns') this.defaultNs = attribute.value;
			else if (attribute.namespaceURI === XMLNS || attribute.name.startsWith('xmlns:'))
				this.prefixes.set(attribute.value, attribute.localName);
		}
	}

	private add(start: number, end: number, text: string): void {
		this.edits.push({ start, end, text, seq: this.seq++ });
	}

	private span(element: XmlElement): Span {
		const span = this.spans.get(element);
		if (!span) throw new Error('Element is not a child of the patched root');
		return span;
	}

	/** The prefix bound to `ns` on the root, declaring `preferred` for it when none is. */
	public prefix(ns: string, preferred: string): string {
		const known = this.prefixes.get(ns);
		if (known !== undefined) return known;
		const taken = new Set(this.prefixes.values());
		let prefix = preferred;
		for (let n = 1; taken.has(prefix); n++) prefix = `${preferred}${n}`;
		this.prefixes.set(ns, prefix);
		const at = this.root.openEnd - (this.root.selfClosing ? 2 : 1);
		this.add(at, at, ` xmlns:${prefix}="${ns}"`);
		return prefix;
	}

	/** `local` qualified for an element in `ns`: unprefixed in the default namespace. */
	public qualify(ns: string, local: string, preferred: string): string {
		return this.defaultNs === ns && !this.prefixes.has(ns)
			? local
			: `${this.prefix(ns, preferred)}:${local}`;
	}

	/** Replaces the element's content, keeping its start tag (and attributes) as written. */
	public setContent(element: XmlElement, content: string): void {
		const span = this.span(element);
		if (span.selfClosing)
			this.add(span.openEnd - 2, span.openEnd, `>${content}</${qualifiedName(this.xml, span)}>`);
		else this.add(span.openEnd, span.closeStart, content);
	}

	/** Adds ` name="value"` text to the element's start tag. */
	public addAttribute(element: XmlElement, attribute: string): void {
		const span = this.span(element);
		const at = span.openEnd - (span.selfClosing ? 2 : 1);
		this.add(at, at, attribute);
	}

	public remove(element: XmlElement): void {
		const span = this.span(element);
		this.add(span.start, span.end, '');
	}

	/** Inserts `markup` before `before`, or as the root's last child. */
	public insert(markup: string, before?: XmlElement): void {
		if (before) {
			const at = this.span(before).start;
			this.add(at, at, markup);
		} else this.appended.push(markup);
	}

	public toString(): string {
		const edits = [...this.edits];
		if (this.appended.length) {
			const text = this.appended.join('');
			const root = this.root;
			if (root.selfClosing) {
				const name = qualifiedName(this.xml, root);
				edits.push({
					start: root.end - 2,
					end: root.end,
					text: `>${text}</${name}>`,
					seq: this.seq,
				});
			} else edits.push({ start: root.closeStart, end: root.closeStart, text, seq: this.seq });
		}
		// Ties: insertions at a position go before a replacement starting there, in call order.
		edits.sort((a, b) => a.start - b.start || a.end - a.start - (b.end - b.start) || a.seq - b.seq);
		let out = '';
		let at = 0;
		for (const edit of edits) {
			out += this.xml.slice(at, edit.start) + edit.text;
			at = Math.max(at, edit.end);
		}
		return out + this.xml.slice(at);
	}
}
