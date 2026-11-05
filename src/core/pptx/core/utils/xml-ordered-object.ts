/**
 * An XML string as the object tree the runtime's XMLBuilder writes, with the
 * document order of every element kept. fast-xml-parser's collapsed shape puts
 * same-tag siblings in one array and orders distinct tags by first occurrence,
 * so an interleaved sequence (`c:barChart`, `c:lineChart`, `c:barChart` in a
 * combination chart, `a:r`, `a:br`, `a:r` in a paragraph) would be written
 * back regrouped. Here a tag that comes back after another one gets the
 * `#pptx-order-N` key marker (`orderedXmlKey`), which the builder strips, so
 * object insertion order carries the sequence. Adjacent repeats stay one array,
 * as the collapsed parser gives them.
 *
 * Text and attribute values are decoded as the runtime parser decodes them
 * (`decodeXmlEntities`) and kept untrimmed: the input is serializer output
 * with no formatting whitespace.
 *
 * @module utils/xml-ordered-object
 */
import { XMLParser } from 'fast-xml-parser';

import { orderedXmlKey } from '../geometry';
import type { XmlObject } from '../types';
import { decodeXmlEntities } from './xml-entities';

type OrderedNode = Record<string, unknown>;

const ATTRIBUTES = ':@';
const TEXT = '#text';

const parser = new XMLParser({
	preserveOrder: true,
	ignoreAttributes: false,
	attributeNamePrefix: '@_',
	parseTagValue: false,
	parseAttributeValue: false,
	trimValues: false,
	processEntities: false,
	tagValueProcessor: (_name: string, value: string) => decodeXmlEntities(value),
	attributeValueProcessor: (_name: string, value: string) => decodeXmlEntities(value),
});

const tagOf = (node: OrderedNode): string | undefined =>
	Object.keys(node).find((key) => key !== ATTRIBUTES);

/** The collapsed value of one element: text, an object of children and attributes, or `''`. */
function valueOf(children: OrderedNode[], attributes: XmlObject | undefined): unknown {
	const text = children.filter((child) => tagOf(child) === TEXT);
	const elements = children.filter((child) => tagOf(child) !== TEXT);
	if (elements.length === 0 && !attributes)
		return text.map((child) => String(child[TEXT])).join('');
	const out: XmlObject = { ...attributes };
	if (text.length > 0) out[TEXT] = text.map((child) => String(child[TEXT])).join('');
	Object.assign(out, childrenOf(elements));
	return out;
}

function childrenOf(nodes: OrderedNode[]): XmlObject {
	const out: XmlObject = {};
	const seen = new Set<string>();
	let previous: string | undefined;
	let key = '';
	nodes.forEach((node, position) => {
		const tag = tagOf(node)!;
		const value = valueOf(
			(node[tag] as OrderedNode[] | undefined) ?? [],
			node[ATTRIBUTES] as XmlObject | undefined,
		);
		if (tag === previous) {
			const current = out[key];
			out[key] = (Array.isArray(current) ? [...current, value] : [current, value]) as XmlObject[];
			return;
		}
		key = seen.has(tag) ? orderedXmlKey(tag, position) : tag;
		seen.add(tag);
		previous = tag;
		Object.defineProperty(out, key, {
			value,
			enumerable: true,
			writable: true,
			configurable: true,
		});
	});
	return out;
}

/** `xml` as an order-keeping object tree for the runtime's XMLBuilder. */
export function orderedXmlObject(xml: string): XmlObject {
	return childrenOf(parser.parse(xml) as OrderedNode[]);
}
