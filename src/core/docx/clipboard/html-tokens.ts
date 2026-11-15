/**
 * A small, DOM-free tokenizer for the HTML Word puts on the clipboard. Word writes HTML 4 with
 * unquoted attributes, values that wrap across lines, namespaced tags (`o:p`), comments that
 * hide content from browsers (`<!--[if supportFields]>`) and "downlevel-revealed" markers that
 * leave content visible (`<![if !supportLists]>`). It is not a general HTML parser: it reads
 * tags, comments, markers and text in order, which is all the normaliser needs.
 */

export type HtmlToken =
	| { kind: 'text'; text: string }
	| { kind: 'comment'; text: string }
	/** `<![if expr]>` (condition) or `<![endif]>` (condition `null`). */
	| { kind: 'marker'; condition: string | null }
	| { kind: 'open'; name: string; attrs: Map<string, string>; selfClosing: boolean }
	| { kind: 'close'; name: string };

const NAME = /[A-Za-z][\w:.-]*/y;
const ATTRIBUTE = /\s*([^\s"'>/=]+)(?:\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s>]+)))?/y;

function readTag(html: string, start: number): { token: HtmlToken; end: number } | null {
	const closing = html[start + 1] === '/';
	NAME.lastIndex = start + (closing ? 2 : 1);
	const name = NAME.exec(html)?.[0];
	if (!name) return null;
	let index = NAME.lastIndex;
	const attrs = new Map<string, string>();
	for (;;) {
		while (/\s/.test(html[index] ?? '')) index++;
		const char = html[index];
		if (char === undefined) return null;
		if (char === '>') {
			const lower = name.toLowerCase();
			return closing
				? { token: { kind: 'close', name: lower }, end: index + 1 }
				: { token: { kind: 'open', name: lower, attrs, selfClosing: false }, end: index + 1 };
		}
		if (char === '/' && html[index + 1] === '>') {
			return {
				token: { kind: 'open', name: name.toLowerCase(), attrs, selfClosing: true },
				end: index + 2,
			};
		}
		ATTRIBUTE.lastIndex = index;
		const match = ATTRIBUTE.exec(html);
		if (!match || ATTRIBUTE.lastIndex === index) {
			index++;
			continue;
		}
		index = ATTRIBUTE.lastIndex;
		const key = match[1]!.toLowerCase();
		if (!attrs.has(key)) attrs.set(key, decodeEntities(match[2] ?? match[3] ?? match[4] ?? ''));
	}
}

/** Splits `html` into tokens; unterminated constructs become text, as a browser shows them. */
export function tokenizeHtml(html: string): HtmlToken[] {
	const tokens: HtmlToken[] = [];
	let text = '';
	const flush = () => {
		if (text) tokens.push({ kind: 'text', text });
		text = '';
	};
	let index = 0;
	while (index < html.length) {
		const lt = html.indexOf('<', index);
		if (lt < 0) {
			text += html.slice(index);
			break;
		}
		text += html.slice(index, lt);
		index = lt;
		if (html.startsWith('<!--', lt)) {
			const end = html.indexOf('-->', lt + 4);
			flush();
			tokens.push({ kind: 'comment', text: html.slice(lt + 4, end < 0 ? html.length : end) });
			index = end < 0 ? html.length : end + 3;
			continue;
		}
		if (html.startsWith('<![', lt)) {
			const end = html.indexOf(']>', lt + 3);
			if (end >= 0) {
				flush();
				const body = html.slice(lt + 3, end).trim();
				const condition = /^if\s+(.*)$/is.exec(body)?.[1]?.trim() ?? null;
				tokens.push({ kind: 'marker', condition: /^endif$/i.test(body) ? null : condition });
				index = end + 2;
				continue;
			}
		}
		if (html[lt + 1] === '!' || html[lt + 1] === '?') {
			// A doctype or processing instruction carries nothing to paste.
			const end = html.indexOf('>', lt);
			index = end < 0 ? html.length : end + 1;
			continue;
		}
		const tag = readTag(html, lt);
		if (!tag) {
			text += '<';
			index = lt + 1;
			continue;
		}
		flush();
		tokens.push(tag.token);
		index = tag.end;
	}
	flush();
	return tokens;
}

/** The named references Word writes, as code points (written as numbers so none is invisible). */
const NAMED: Record<string, number> = {
	amp: 0x26,
	lt: 0x3c,
	gt: 0x3e,
	quot: 0x22,
	apos: 0x27,
	nbsp: 0xa0,
	shy: 0xad,
	ensp: 0x2002,
	emsp: 0x2003,
	ndash: 0x2013,
	mdash: 0x2014,
	lsquo: 0x2018,
	rsquo: 0x2019,
	ldquo: 0x201c,
	rdquo: 0x201d,
	bull: 0x2022,
	hellip: 0x2026,
	middot: 0xb7,
	copy: 0xa9,
	reg: 0xae,
	trade: 0x2122,
};

/** Decodes character references; unknown names stay as written. */
export function decodeEntities(value: string): string {
	return value.replace(/&(#x[0-9a-f]+|#\d+|[a-z]+\d*);/gi, (whole, body: string) => {
		if (body[0] === '#') {
			const code =
				body[1] === 'x' || body[1] === 'X' ? parseInt(body.slice(2), 16) : Number(body.slice(1));
			return Number.isInteger(code) && code > 0 && code <= 0x10ffff
				? String.fromCodePoint(code)
				: whole;
		}
		const named = NAMED[body.toLowerCase()];
		return named === undefined ? whole : String.fromCodePoint(named);
	});
}

export const escapeText = (value: string): string =>
	value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

export const escapeAttribute = (value: string): string => escapeText(value).replace(/"/g, '&quot;');

/** Serialises an element's start tag from an attribute record (empty values are dropped). */
export function startTag(name: string, attrs: Record<string, string | undefined>): string {
	const parts = Object.entries(attrs)
		.filter((entry): entry is [string, string] => Boolean(entry[1]))
		.map(([key, value]) => ` ${key}="${escapeAttribute(value)}"`);
	return `<${name}${parts.join('')}>`;
}

/** The text of `tokens`, ignoring markup (used for field codes and list markers). */
export function tokensText(tokens: readonly HtmlToken[]): string {
	return tokens.map((token) => (token.kind === 'text' ? decodeEntities(token.text) : '')).join('');
}
