import { elements, isElement, parseXml, type XmlElement } from '../../xml/xml';
import type { InkXmlObject } from './types';

/**
 * The InkML object tree read from, and written back to, part text through the shared XML parser,
 * so Word and Excel (which parse with the DOM) feed the same reader and writer PowerPoint feeds
 * from its own parser. The shape matches that parser's: attributes as `@_name`, text as `#text`,
 * repeated children as arrays and an element holding only text as a bare string.
 *
 * @module drawingml/inkml/xml-tree
 */

const DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

function textOf(element: XmlElement): string {
	let text = '';
	for (let child = element.firstChild; child; child = child.nextSibling)
		if (child.nodeType === 3 || child.nodeType === 4) text += child.nodeValue ?? '';
	return text;
}

function convert(element: XmlElement): InkXmlObject | string {
	const object: InkXmlObject = {};
	for (let index = 0; index < element.attributes.length; index++) {
		const attribute = element.attributes.item(index);
		if (attribute) object[`@_${attribute.name}`] = attribute.value;
	}
	const children = elements(element);
	const text = textOf(element);
	if (!children.length && element.attributes.length === 0) return text;
	if (!children.length && text) object['#text'] = text;
	for (const child of children) {
		const value = convert(child);
		const existing = object[child.nodeName];
		if (existing === undefined) object[child.nodeName] = value as InkXmlObject;
		else if (Array.isArray(existing)) existing.push(value as InkXmlObject);
		else object[child.nodeName] = [existing as InkXmlObject, value as InkXmlObject];
	}
	return object;
}

/** Parses InkML part text into the object tree; throws on malformed XML, like every part reader. */
export function inkObjectFromXml(xml: string): InkXmlObject {
	const root = parseXml(xml, { label: 'InkML', limits: { maxDepth: 64 } }).documentElement;
	if (!root || !isElement(root)) return {};
	const tree: InkXmlObject = {};
	tree[root.nodeName] = convert(root) as InkXmlObject;
	return tree;
}

const escapeText = (value: string): string =>
	value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const escapeAttribute = (value: string): string => escapeText(value).replace(/"/g, '&quot;');

function serialize(name: string, value: InkXmlObject | string, out: string[]): void {
	if (typeof value === 'string') {
		out.push(value ? `<${name}>${escapeText(value)}</${name}>` : `<${name}/>`);
		return;
	}
	let open = `<${name}`;
	const body: string[] = [];
	for (const [key, child] of Object.entries(value)) {
		if (child === undefined) continue;
		if (key.startsWith('@_')) open += ` ${key.slice(2)}="${escapeAttribute(String(child))}"`;
		else if (key === '#text') body.push(escapeText(String(child)));
		else if (Array.isArray(child)) for (const item of child) serialize(key, item, body);
		else serialize(key, child as InkXmlObject | string, body);
	}
	out.push(body.length ? `${open}>${body.join('')}</${name}>` : `${open}/>`);
}

/** Serializes an object tree (one root element) as a standalone XML part. */
export function inkObjectToXml(tree: InkXmlObject): string {
	const out: string[] = [DECLARATION];
	for (const [key, value] of Object.entries(tree)) {
		if (key.startsWith('?') || key.startsWith('@_') || value === undefined) continue;
		if (Array.isArray(value)) for (const item of value) serialize(key, item, out);
		else serialize(key, value as InkXmlObject | string, out);
	}
	return out.join('');
}
