import type { ChartXmlSource } from '../../../chart/keep-source';
import { NS, isElement, withoutByteOrderMark, type XmlElement } from '../../../xml/index';
import { annotateParagraphSiblingOrder } from '../core/runtime/paragraph-sibling-order';
import type { XmlObject } from '../types';
import { preservesXmlWhitespace } from './xml-whitespace';

// The stock PPTX parser keeps attributes and text as strings, groups children by their
// qualified name, and collapses text-only elements. Custom parsers keep their own behaviour.
interface CompatibilityParser {
	parse(text: string): unknown;
}

const standardParsers = new WeakMap<
	CompatibilityParser,
	{ parse: CompatibilityParser['parse']; maxDepth: number }
>();

export function registerChartCompatibilityParser(
	parser: CompatibilityParser,
	maxDepth: number,
): void {
	standardParsers.set(parser, { parse: parser.parse, maxDepth });
}

function project(
	element: XmlElement,
	mayNormalizeAttributes: boolean,
	remainingDepth: number,
): XmlObject | string | undefined {
	if (remainingDepth <= 0) return undefined;
	let object: XmlObject | undefined;
	let text = '';
	for (let child = element.firstChild; child; child = child.nextSibling) {
		if (!isElement(child)) {
			if (child.nodeType !== 3) return undefined;
			const value = child.nodeValue ?? '';
			text += preservesXmlWhitespace(element.tagName) ? value : value.trim();
			continue;
		}
		const name = child.tagName;
		// The compatibility parser rejects or renames reserved JS keys.
		if (name in Object.prototype || name === 'prototype') return undefined;
		const value = project(child, mayNormalizeAttributes, remainingDepth - 1);
		if (value === undefined) return undefined;
		object ??= {};
		const previous = object[name];
		if (previous === undefined) object[name] = value;
		else if (typeof previous === 'string' || typeof value === 'string') return undefined;
		else if (Array.isArray(previous)) previous.push(value);
		else object[name] = [previous, value];
	}
	if (text && (object || element.attributes.length)) (object ??= {})['#text'] = text;
	for (let i = 0; i < element.attributes.length; i++) {
		const attribute = element.attributes.item(i);
		if (!attribute) continue;
		// XML normalizes literal tabs and newlines inside attributes to spaces. The object
		// parser only normalizes their line endings. Keep it when a value may differ.
		if (mayNormalizeAttributes && attribute.value.includes(' ')) return undefined;
		(object ??= {})[`@_${attribute.name}`] = attribute.value;
	}
	return object ?? text;
}

/** The read-only compatibility tree of an ordinary chart, from its already parsed DOM. */
export function projectChartCompatibilityTree(
	parser: CompatibilityParser,
	text: string,
	part: ChartXmlSource | undefined,
): XmlObject | undefined {
	const standard = standardParsers.get(parser);
	if (
		!part ||
		part.hadWarnings ||
		part.root.localName !== 'chartSpace' ||
		(part.root.namespaceURI && part.root.namespaceURI !== NS.c) ||
		standard?.parse !== parser.parse ||
		// Keep the compatibility parser for its special character and sibling-order handling.
		/&#|[\u0085\u2028\u2029]|custGeom|dataModel|txBody|oMath/u.test(text)
	)
		return undefined;
	const source = withoutByteOrderMark(text);
	const closingDeclaration = source.startsWith('<?xml') ? source.indexOf('?>') + 2 : 0;
	const body = source.slice(closingDeclaration).trimStart();
	const root = part.root;
	// Other top-level markup (including PIs) may influence the object parser.
	if (!body.startsWith(`<${root.tagName}`)) return undefined;
	const tree = project(root, /[\t\r\n]/u.test(body), standard.maxDepth);
	if (tree === undefined) return undefined;
	const projected: XmlObject = {};
	projected[root.tagName] = tree;
	if (text.includes('<a:p')) annotateParagraphSiblingOrder(text, projected);
	return projected;
}
