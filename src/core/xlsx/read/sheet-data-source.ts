// A worksheet's `<sheetData>` holds nearly all of its XML (every cell). Building a full DOM for it
// dominated opening a large workbook, so the cell data is cut out of the part and read with the
// lightweight fragment reader, one row at a time so that only the current row exists as elements;
// the rest of the part (views, columns, formatting, ...) still goes through the DOM. Only plain
// markup takes this path, with element prefixes declared on the root: comments and CDATA sections
// may appear anywhere, but nested namespace declarations, processing instructions in the cell data
// or anything else the light reader declines is parsed whole, as before, so the result is the same
// either way. The strings of the shared strings part are read the same way (`rootChildren`).
import { liteChildElements, type LiteNamespaces, type XmlElement } from '../../xml/index';

export interface SplitSheetData {
	/** The part with `<sheetData>` emptied, for the DOM parser. */
	xml: string;
	/**
	 * The children of `<sheetData>` (its rows), readable through the same DOM subset as parsed
	 * ones and built one at a time while iterating. Markup the light reader declines throws a
	 * `LiteDeclinedError` from the iteration, possibly after some rows: the part must then be
	 * read whole.
	 */
	rows: Iterable<XmlElement>;
}

const ROOT = /^\uFEFF?(?:\s*<\?[^]*?\?>)?\s*<([A-Za-z_][\w.:-]*)((?:\s[^>]*)?)>/;
const DECLARATION = /\sxmlns(?::([\w.-]+))?\s*=\s*(["'])(.*?)\2/g;
const OPEN = /<((?:[A-Za-z_][\w.-]*:)?)sheetData(?=[\s/>])/;

/** The namespaces declared on the root element, or undefined when it cannot be read. */
function rootNamespaces(xml: string): LiteNamespaces | undefined {
	const root = ROOT.exec(xml.slice(0, 8192));
	if (!root) return undefined;
	const out: Record<string, string> = {};
	for (const match of (root[2] ?? '').matchAll(DECLARATION)) out[match[1] ?? ''] = match[3] ?? '';
	return out;
}

const SPECIAL = /<[!?]/g;
const SECTIONS: readonly (readonly [open: string, close: string])[] = [
	['<!--', '-->'],
	['<![CDATA[', ']]>'],
	['<?', '?>'],
];

/**
 * `xml` with its comments and CDATA sections replaced by spaces of the same length, so a search for
 * tags never matches text inside them and every position still points into `xml`; undefined for a
 * declaration (`<!DOCTYPE`, ...) or an unterminated section.
 */
function blankComments(xml: string): string | undefined {
	if (!xml.includes('<!--') && !xml.includes('<![CDATA[')) return xml;
	const pieces: string[] = [];
	let copied = 0;
	SPECIAL.lastIndex = 0;
	for (let match = SPECIAL.exec(xml); match; match = SPECIAL.exec(xml)) {
		const at = match.index;
		const section = SECTIONS.find(([open]) => xml.startsWith(open, at));
		if (!section) return undefined;
		const [open, close] = section;
		const end = xml.indexOf(close, at + open.length);
		if (end < 0) return undefined;
		const after = end + close.length;
		// Processing instructions are only stepped over: the root's declaration stays readable.
		if (open !== '<?') {
			pieces.push(xml.slice(copied, at), ' '.repeat(after - at));
			copied = after;
		}
		SPECIAL.lastIndex = after;
	}
	pieces.push(xml.slice(copied));
	return pieces.join('');
}

/** `parseXml` refuses a part holding these anywhere, even in a comment outside the root. */
const DECLARATIONS = /<!DOCTYPE|<!ENTITY/i;

/**
 * The children of a part's root element (the strings of the shared strings part), read by the
 * light reader one at a time like the rows of {@link SplitSheetData}, with the namespaces declared
 * on the root; undefined when the part must go to the DOM parser (a declaration, a self-closing
 * root, anything but whitespace, comments and an XML declaration around the root).
 */
export function rootChildren(xml: string): Iterable<XmlElement> | undefined {
	if (DECLARATIONS.test(xml)) return undefined;
	const plain = blankComments(xml);
	if (plain === undefined) return undefined;
	const namespaces = rootNamespaces(plain);
	const root = ROOT.exec(plain.slice(0, 8192));
	if (!namespaces || !root || (root[2] ?? '').trimEnd().endsWith('/')) return undefined;
	const close = plain.lastIndexOf(`</${root[1] ?? ''}`);
	if (close < root[0].length || !/^<\/[^\s>]+\s*>\s*$/.test(plain.slice(close))) return undefined;
	const content = xml.slice(root[0].length, close);
	return { [Symbol.iterator]: () => liteChildElements(content, namespaces) };
}

/** Splits the cell data from a worksheet part, or undefined when the whole part must be parsed. */
export function splitSheetData(xml: string): SplitSheetData | undefined {
	// Tags are found in `plain` (comments and CDATA blanked); the parts are cut from `xml`.
	const plain = blankComments(xml);
	if (plain === undefined) return undefined;
	const namespaces = rootNamespaces(plain);
	const open = OPEN.exec(plain);
	if (!namespaces || !open) return undefined;
	const prefix = (open[1] ?? '').slice(0, -1);
	const ns = namespaces[prefix];
	if (prefix && ns === undefined) return undefined;
	const tagName = `${open[1] ?? ''}sheetData`;
	const openEnd = plain.indexOf('>', open.index);
	if (openEnd < 0) return undefined;
	const attributes = plain.slice(open.index + tagName.length + 1, openEnd);
	if (attributes.trimEnd().endsWith('/') || attributes.includes('xmlns')) return undefined;
	const close = plain.indexOf(`</${tagName}`, openEnd);
	if (close < 0 || OPEN.test(plain.slice(openEnd))) return undefined;
	const data = xml.slice(openEnd + 1, close);
	return {
		xml: xml.slice(0, openEnd + 1) + xml.slice(close),
		rows: { [Symbol.iterator]: () => liteChildElements(data, namespaces) },
	};
}
