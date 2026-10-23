// Shared plumbing for the DrawingML text writers: element creation in the `a` namespace, verbatim
// children re-imported from their XML, and schema-ordered child placement so kept children land
// where the schema puts them.
import { NS, elements, parseXml, type XmlDocument, type XmlElement } from '../xml/index';

/** A new `a:<local>` element. */
export function drawingElement(doc: XmlDocument, local: string): XmlElement {
	return doc.createElementNS(NS.a, `a:${local}`);
}

/** Removes declarations of namespaces already in scope at the target (`a`, `r`). */
function dropRedundantDeclarations(element: XmlElement): void {
	if (element.getAttribute('xmlns:a') === NS.a) element.removeAttribute('xmlns:a');
	if (element.getAttribute('xmlns:r') === NS.r) element.removeAttribute('xmlns:r');
	for (const child of elements(element)) dropRedundantDeclarations(child);
}

/** Imports one element written as XML (a kept child, or a fragment from another writer). */
export function importDrawingXml(doc: XmlDocument, xml: string): XmlElement {
	const wrapper = parseXml(`<w xmlns:a="${NS.a}" xmlns:r="${NS.r}">${xml}</w>`).documentElement;
	const node = elements(wrapper)[0];
	if (!node) throw new Error('Empty DrawingML fragment');
	const imported = doc.importNode(node, true) as XmlElement;
	dropRedundantDeclarations(imported);
	return imported;
}

/** One child to place, by local name. */
export interface OrderedChild {
	name: string;
	node: XmlElement;
}

/** The local name of a kept child's XML (`<a:effectLst ...>` is `effectLst`). */
export function keptChildName(xml: string): string {
	const match = /^\s*<(?:[\w.-]+:)?([\w.-]+)/u.exec(xml);
	return match?.[1] ?? '';
}

/** Kept children as ordered entries. */
export function keptChildren(doc: XmlDocument, xml: string[] | undefined): OrderedChild[] {
	return (xml ?? []).map((value) => ({
		name: keptChildName(value),
		node: importDrawingXml(doc, value),
	}));
}

/**
 * Appends `entries` to `parent` in schema order: by the position of each name in `order`, names
 * the schema does not list going last, ties keeping their given order.
 */
export function appendInSchemaOrder(
	parent: XmlElement,
	order: readonly string[],
	entries: OrderedChild[],
): void {
	const rank = (name: string) => {
		const index = order.indexOf(name);
		return index < 0 ? order.length : index;
	};
	entries
		.map((entry, index) => ({ entry, index }))
		.sort((a, b) => rank(a.entry.name) - rank(b.entry.name) || a.index - b.index)
		.forEach(({ entry }) => parent.appendChild(entry.node));
}

/** Writes `'1'`/`'0'` for a defined boolean. */
export function setBoolean(element: XmlElement, name: string, value: boolean | undefined): void {
	if (value !== undefined) element.setAttribute(name, value ? '1' : '0');
}

/** Writes a defined value as a string. */
export function setValue(
	element: XmlElement,
	name: string,
	value: string | number | undefined,
): void {
	if (value !== undefined) element.setAttribute(name, String(value));
}

/** Writes a defined number rounded to an integer after scaling. */
export function setScaled(
	element: XmlElement,
	name: string,
	value: number | undefined,
	scale: number,
): void {
	if (value !== undefined && Number.isFinite(value))
		element.setAttribute(name, String(Math.round(value * scale)));
}
