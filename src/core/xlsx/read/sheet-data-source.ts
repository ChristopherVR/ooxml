// A worksheet's `<sheetData>` holds nearly all of its XML (every cell). Building a full DOM for it
// dominated opening a large workbook, so the cell data is cut out of the part and read with the
// lightweight fragment reader; the rest of the part (views, columns, formatting, ...) still goes
// through the DOM. Only plain markup takes this path, with element prefixes declared on the root:
// comments, CDATA, nested namespace declarations or anything else the light reader declines is
// parsed whole, as before, so the result is the same either way.
import {
	LiteElement,
	parseLiteFragment,
	type LiteNamespaces,
	type XmlElement,
} from '../../xml/index';

export interface SplitSheetData {
	/** The part with `<sheetData>` emptied, for the DOM parser. */
	xml: string;
	/** The cell data, readable through the same DOM subset as a parsed `<sheetData>`. */
	sheetData: XmlElement;
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

/** Splits the cell data from a worksheet part, or undefined when the whole part must be parsed. */
export function splitSheetData(xml: string): SplitSheetData | undefined {
	if (xml.includes('<!--') || xml.includes('<![CDATA[')) return undefined;
	const namespaces = rootNamespaces(xml);
	const open = OPEN.exec(xml);
	if (!namespaces || !open) return undefined;
	const prefix = (open[1] ?? '').slice(0, -1);
	const ns = namespaces[prefix];
	if (prefix && ns === undefined) return undefined;
	const tagName = `${open[1] ?? ''}sheetData`;
	const openEnd = xml.indexOf('>', open.index);
	if (openEnd < 0) return undefined;
	const attributes = xml.slice(open.index + tagName.length + 1, openEnd);
	if (attributes.trimEnd().endsWith('/') || attributes.includes('xmlns')) return undefined;
	const close = xml.indexOf(`</${tagName}`, openEnd);
	if (close < 0 || OPEN.test(xml.slice(openEnd))) return undefined;
	const container = new LiteElement(tagName, 'sheetData', ns || null);
	const data = parseLiteFragment(xml.slice(openEnd + 1, close), container, namespaces);
	if (!data) return undefined;
	return { xml: xml.slice(0, openEnd + 1) + xml.slice(close), sheetData: data.asElement() };
}
