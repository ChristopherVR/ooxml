/**
 * Bounded, DOM-free HTML lexing for clipboard markup (`text/html`, the Windows `HTML Format`
 * (CF_HTML) payload). Nothing here executes or resolves anything: script and style content is
 * skipped as raw text, entities are decoded only to valid code points, and every scan is linear.
 *
 * The entity, attribute, CSS declaration and CSS colour helpers were first written for Excel's
 * table paste (`xlsx/edit/clipboard-html-parse.ts`), which now imports them from here.
 */

const ENTITIES: Record<string, string> = {
	amp: '&',
	lt: '<',
	gt: '>',
	quot: '"',
	apos: "'",
	nbsp: '\u00a0',
	ensp: '\u2002',
	emsp: '\u2003',
	thinsp: '\u2009',
	ndash: '\u2013',
	mdash: '\u2014',
	lsquo: '\u2018',
	rsquo: '\u2019',
	ldquo: '\u201c',
	rdquo: '\u201d',
	bull: '\u2022',
	hellip: '\u2026',
	copy: '\u00a9',
	reg: '\u00ae',
	trade: '\u2122',
	euro: '\u20ac',
};

/** A numeric reference as text, or U+FFFD for NUL, surrogates and values past U+10FFFF. */
function codePoint(value: number): string {
	if (!Number.isFinite(value) || value <= 0 || value > 0x10ffff) return '\ufffd';
	if (value >= 0xd800 && value <= 0xdfff) return '\ufffd';
	return String.fromCodePoint(value);
}

/** Decodes `&name;`, `&#n;` and `&#xh;` references; unknown names are left as written. */
export function decodeHtmlEntities(text: string): string {
	return text.replace(/&(#x[0-9a-f]{1,8}|#\d{1,10}|[a-z]{1,10});/gi, (whole, name: string) => {
		if (name[1] === 'x' || name[1] === 'X') return codePoint(parseInt(name.slice(2), 16));
		if (name.startsWith('#')) return codePoint(Number(name.slice(1)));
		return ENTITIES[name.toLowerCase()] ?? whole;
	});
}

/** The attributes of a start tag's inside (`a href="x" title='y'`), names lower-cased. */
export function parseHtmlAttributes(text: string): Map<string, string> {
	const attrs = new Map<string, string>();
	for (const m of text.matchAll(
		/([\w:-]{1,64})\s{0,16}=\s{0,16}(?:"([^"]*)"|'([^']*)'|([^\s>]+))/g,
	))
		attrs.set((m[1] ?? '').toLowerCase(), decodeHtmlEntities(m[2] ?? m[3] ?? m[4] ?? ''));
	return attrs;
}

/** A `style` attribute's declarations, property names lower-cased. */
export function parseCssDeclarations(text: string): Map<string, string> {
	const out = new Map<string, string>();
	for (const decl of text.split(';')) {
		const colon = decl.indexOf(':');
		if (colon < 0) continue;
		out.set(
			decl.slice(0, colon).trim().toLowerCase(),
			decl
				.slice(colon + 1)
				.replace(/!important/gi, '')
				.trim(),
		);
	}
	return out;
}

const NAMED_COLORS: Record<string, string> = {
	black: '000000',
	white: 'FFFFFF',
	red: 'FF0000',
	green: '008000',
	lime: '00FF00',
	blue: '0000FF',
	yellow: 'FFFF00',
	orange: 'FFA500',
	purple: '800080',
	gray: '808080',
	grey: '808080',
	silver: 'C0C0C0',
	navy: '000080',
	maroon: '800000',
	teal: '008080',
	olive: '808000',
	aqua: '00FFFF',
	cyan: '00FFFF',
	fuchsia: 'FF00FF',
	magenta: 'FF00FF',
	darkblue: '00008B',
	darkred: '8B0000',
	darkgreen: '006400',
	darkcyan: '008B8B',
	darkmagenta: '8B008B',
	darkgray: 'A9A9A9',
	darkgrey: 'A9A9A9',
	lightgray: 'D3D3D3',
	lightgrey: 'D3D3D3',
	darkyellow: '808000',
};

/** A CSS colour (`#rgb`, `#rrggbb`, `rgb()`, a basic name) as upper-case `RRGGBB`. */
export function parseCssColorHex(text: string): string | undefined {
	const value = text.trim().toLowerCase();
	const hex = /^#([0-9a-f]{3}|[0-9a-f]{6})$/.exec(value)?.[1];
	if (hex) return (hex.length === 3 ? [...hex].map((ch) => ch + ch).join('') : hex).toUpperCase();
	const rgb = /^rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)(?:\s*,\s*([\d.]+))?/.exec(value);
	if (rgb) {
		// A fully transparent colour paints nothing (browsers copy `rgba(0, 0, 0, 0)` backgrounds).
		if (rgb[4] !== undefined && Number(rgb[4]) === 0) return undefined;
		return rgb
			.slice(1, 4)
			.map((n) => Math.min(255, Number(n)).toString(16).padStart(2, '0'))
			.join('')
			.toUpperCase();
	}
	return NAMED_COLORS[value];
}

/** One lexical item of clipboard HTML. */
export type HtmlToken =
	| { kind: 'open'; name: string; attrs: Map<string, string>; selfClosing: boolean }
	| { kind: 'close'; name: string }
	| { kind: 'text'; text: string }
	/** Word's `<![if !supportLists]>` (true) and `<![endif]>` (false) conditional markers. */
	| { kind: 'conditional'; open: boolean; supportLists: boolean };

/** Elements whose content is never markup and never pasted text. */
const MAX_TAG_LENGTH = 8192;

const RAW_TEXT = new Set([
	'script',
	'style',
	'textarea',
	'title',
	'xmp',
	'noscript',
	'template',
	'head',
	'xml',
	'svg',
	'math',
	'iframe',
	'object',
	'canvas',
	'select',
]);

/** The fragment of a CF_HTML or browser payload between its `StartFragment` / `EndFragment`. */
export function htmlClipboardFragment(html: string): string {
	const start = html.indexOf('<!--StartFragment-->');
	const end = html.indexOf('<!--EndFragment-->', start < 0 ? 0 : start);
	if (start >= 0 && end > start) return html.slice(start + '<!--StartFragment-->'.length, end);
	// A CF_HTML header without fragment comments still starts with `Version:` lines.
	const first = html.indexOf('<');
	return first > 0 && /^Version:\d/.test(html) ? html.slice(first) : html;
}

/** Index of the `>` closing a tag opened at `from`, honouring quoted attribute values. */
function tagEnd(html: string, from: number): number {
	let quote = '';
	// A start tag longer than this is not markup a clipboard writes; the cap keeps the scan linear.
	const limit = Math.min(html.length, from + MAX_TAG_LENGTH);
	for (let i = from; i < limit; i++) {
		const ch = html[i];
		if (quote) {
			if (ch === quote) quote = '';
		} else if (ch === '"' || ch === "'") quote = ch;
		else if (ch === '>') return i;
		else if (ch === '<') return -1;
	}
	return -1;
}

/**
 * Splits HTML into tokens in one pass. Comments, doctypes and processing instructions are
 * dropped, raw-text elements (`script`, `style`...) are skipped with their content, and a `<`
 * that does not start a tag stays text. At most `maxTokens` tokens are produced; the rest of the
 * input is then returned as one text token with its tags removed.
 */
export function tokenizeHtml(html: string, maxTokens = Number.POSITIVE_INFINITY): HtmlToken[] {
	const tokens: HtmlToken[] = [];
	let text = '';
	const flush = (): void => {
		if (text) tokens.push({ kind: 'text', text: decodeHtmlEntities(text) });
		text = '';
	};
	let i = 0;
	while (i < html.length) {
		if (tokens.length >= maxTokens) {
			text += textWithoutTags(html.slice(i));
			break;
		}
		const lt = html.indexOf('<', i);
		if (lt < 0) {
			text += html.slice(i);
			break;
		}
		text += html.slice(i, lt);
		i = lt;
		if (html.startsWith('<!--', i)) {
			const close = html.indexOf('-->', i + 4);
			i = close < 0 ? html.length : close + 3;
			continue;
		}
		if (html.startsWith('<![', i)) {
			const close = html.indexOf(']>', i + 3);
			const body = html.slice(i + 3, close < 0 ? html.length : close).toLowerCase();
			flush();
			if (/^\s*if\b/.test(body))
				tokens.push({
					kind: 'conditional',
					open: true,
					supportLists: body.includes('supportlists'),
				});
			else if (/^\s*endif\b/.test(body))
				tokens.push({ kind: 'conditional', open: false, supportLists: false });
			i = close < 0 ? html.length : close + 2;
			continue;
		}
		const next = html[i + 1] ?? '';
		if (next === '!' || next === '?') {
			const close = html.indexOf('>', i + 2);
			i = close < 0 ? html.length : close + 1;
			continue;
		}
		const closing = next === '/';
		const nameMatch = /^[a-zA-Z][\w:-]{0,63}/.exec(html.slice(i + (closing ? 2 : 1), i + 70));
		const end = nameMatch ? tagEnd(html, i + 1) : -1;
		if (!nameMatch || end < 0) {
			text += '<';
			i += 1;
			continue;
		}
		const name = nameMatch[0].toLowerCase();
		flush();
		if (closing) {
			tokens.push({ kind: 'close', name });
			i = end + 1;
			continue;
		}
		const inner = html.slice(i + 1 + nameMatch[0].length, end);
		const selfClosing = inner.trimEnd().endsWith('/');
		i = end + 1;
		if (RAW_TEXT.has(name)) {
			// Counted as a token so a flood of empty raw elements still meets the cap.
			tokens.push({ kind: 'text', text: '' });
			if (selfClosing) continue;
			const closer = new RegExp(`</${name}`, 'gi');
			closer.lastIndex = i;
			const found = closer.exec(html);
			const after = found ? html.indexOf('>', found.index) : -1;
			i = after < 0 ? html.length : after + 1;
			continue;
		}
		tokens.push({ kind: 'open', name, attrs: parseHtmlAttributes(inner), selfClosing });
	}
	flush();
	return tokens;
}

/**
 * `html` with every complete tag (`<` up to the next `>`) removed, in one linear pass; a `<`
 * with no closing `>` stays as text.
 */
function textWithoutTags(html: string): string {
	let out = '';
	let at = 0;
	for (let lt = html.indexOf('<'); lt >= 0; lt = html.indexOf('<', at)) {
		const gt = html.indexOf('>', lt + 1);
		if (gt < 0) break;
		const inner = html.indexOf('<', lt + 1);
		// `a < b <i>`: the first `<` opens no tag, so it stays text.
		if (inner >= 0 && inner < gt) {
			out += html.slice(at, inner);
			at = inner;
			continue;
		}
		out += html.slice(at, lt);
		at = gt + 1;
	}
	return out + html.slice(at);
}
