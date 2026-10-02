export const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

// Characters XML 1.0 cannot carry, plus a lone CR that parsers would normalise away.
const CONTROL = '\u0000-\u0008\u000B\u000C\u000E-\u001F' + String.fromCharCode(0xfffe, 0xffff);
const NEEDS_ESCAPE = new RegExp(`[${CONTROL}\r]|_x[0-9A-Fa-f]{4}_`);
const ESCAPABLE = new RegExp(`[${CONTROL}\r]|_(?=x[0-9A-Fa-f]{4}_)`, 'g');

/**
 * Encodes text for a SpreadsheetML string: control characters (and `\r`) become `_xHHHH_`, and
 * an underscore that would read as an escape is itself escaped as `_x005F_`.
 */
export function encodeEscapes(text: string): string {
	if (!NEEDS_ESCAPE.test(text)) return text;
	return text.replace(
		ESCAPABLE,
		(ch) => `_x${ch.charCodeAt(0).toString(16).toUpperCase().padStart(4, '0')}_`,
	);
}

const INVALID_XML = new RegExp(`[${CONTROL}]`, 'g');

/** Escapes element text; characters XML cannot carry are dropped. */
export const escapeText = (text: string): string =>
	text.replace(INVALID_XML, '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

/** Escapes an attribute value (including newlines and tabs, which would be normalised). */
export const escapeAttr = (text: string): string =>
	text
		.replace(INVALID_XML, '')
		.replace(/&/g, '&amp;')
		.replace(/</g, '&lt;')
		.replace(/>/g, '&gt;')
		.replace(/"/g, '&quot;')
		.replace(/\r/g, '&#13;')
		.replace(/\n/g, '&#10;')
		.replace(/\t/g, '&#9;');

export type AttrValue = string | number | boolean | undefined;

/** ` a="1" b="x"` for the defined entries; booleans become `1`/`0`. */
export function attrs(values: Record<string, AttrValue>): string {
	let out = '';
	for (const [name, value] of Object.entries(values)) {
		if (value === undefined) continue;
		const text =
			typeof value === 'boolean'
				? value
					? '1'
					: '0'
				: typeof value === 'number'
					? num(value)
					: value;
		out += ` ${name}="${escapeAttr(text)}"`;
	}
	return out;
}

/** An element: self-closing when it has no content. */
export const el = (name: string, values: Record<string, AttrValue> = {}, content = ''): string =>
	content ? `<${name}${attrs(values)}>${content}</${name}>` : `<${name}${attrs(values)}/>`;

/** A number as SpreadsheetML writes it (shortest round-trip form). */
export const num = (value: number): string => (Object.is(value, -0) ? '0' : String(value));

/** `<t>` text with `xml:space="preserve"` when whitespace would otherwise be trimmed. */
export function tElement(text: string, name = 't'): string {
	const encoded = escapeText(encodeEscapes(text));
	const preserve = /^\s|\s$|\n|\t/.test(text) ? ' xml:space="preserve"' : '';
	return `<${name}${preserve}>${encoded}</${name}>`;
}

const MAIN_NS = ' xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';

/**
 * Removes the SpreadsheetML default namespace declaration from a fragment's first tag; kept
 * fragments are serialized standalone and inherit it again from the part's root.
 */
export function inlineFragment(xml: string): string {
	const end = xml.indexOf('>');
	if (end < 0) return xml;
	const head = xml.slice(0, end);
	return head.includes(MAIN_NS) ? head.replace(MAIN_NS, '') + xml.slice(end) : xml;
}
