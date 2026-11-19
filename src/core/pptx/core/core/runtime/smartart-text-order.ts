import type { XmlObject } from '../../types';
import { at } from '../../../../geometry/indexed';
import { xmlTags } from '../../../../xml/tags';

const childOrder = new WeakMap<XmlObject, string[]>();

function localName(name: string): string {
	const colon = name.indexOf(':');
	return colon >= 0 ? name.slice(colon + 1) : name;
}

function isXmlObject(value: unknown): value is XmlObject {
	return Boolean(value && typeof value === 'object' && !Array.isArray(value));
}

function collectParsedParagraphs(root: unknown): XmlObject[] {
	const paragraphs: XmlObject[] = [];
	const stack: unknown[] = [root];
	while (stack.length > 0) {
		const current = stack.pop();
		if (!current || typeof current !== 'object') {
			continue;
		}
		if (Array.isArray(current)) {
			for (let index = current.length - 1; index >= 0; index--) {
				stack.push(current[index]);
			}
			continue;
		}
		const node = current as XmlObject;
		for (const [key, value] of Object.entries(node).reverse()) {
			if (
				(localName(key) === 't' || localName(key) === 'txBody') &&
				value &&
				typeof value === 'object'
			) {
				for (const [childKey, childValue] of Object.entries(value as XmlObject)) {
					if (localName(childKey) !== 'p') {
						continue;
					}
					paragraphs.push(
						...(Array.isArray(childValue) ? childValue : [childValue]).filter(isXmlObject),
					);
				}
			} else {
				stack.push(value);
			}
		}
	}
	return paragraphs;
}

/**
 * The local names of the direct children of every `elementName` element in `xml`, in document
 * order. An empty element parses to an object only when it has attributes, so only then does it
 * get an (empty) entry: entries pair with the parsed objects by index.
 */
function extractElementOrders(xml: string, elementName: string): string[][] {
	const orders: string[][] = [];
	let open: string | undefined;
	let depth = 0;
	let order: string[] = [];
	for (const tag of xmlTags(xml)) {
		if (open === undefined) {
			if (tag.closing || tag.local !== elementName) continue;
			if (tag.selfClosing) {
				if (tag.attributes.trim()) orders.push([]);
				continue;
			}
			open = tag.name;
			depth = 0;
			order = [];
		} else if (tag.closing) {
			if (tag.name === open) {
				orders.push(order);
				open = undefined;
			} else depth = Math.max(0, depth - 1);
		} else {
			if (depth === 0) order.push(tag.local);
			if (!tag.selfClosing) depth++;
		}
	}
	return orders;
}

function directChildrenByLocalName(paragraphs: XmlObject[], name: string): XmlObject[] {
	return paragraphs.flatMap((paragraph) =>
		Object.entries(paragraph).flatMap(([key, value]) => {
			if (localName(key) !== name) {
				return [];
			}
			const values = Array.isArray(value) ? value : [value];
			return values.filter(isXmlObject);
		}),
	);
}

/** Attach source item order to parsed SmartArt paragraph objects. */
export function annotateSmartArtTextOrder(xml: string, parsed: unknown): void {
	const paragraphs = collectParsedParagraphs(parsed);
	const orders = extractElementOrders(xml, 'p');
	for (let index = 0; index < Math.min(orders.length, paragraphs.length); index++) {
		childOrder.set(at(paragraphs, index), at(orders, index));
	}
	for (const name of ['r', 'fld', 'br']) {
		const nodes = directChildrenByLocalName(paragraphs, name);
		const nodeOrders = extractElementOrders(xml, name);
		for (let index = 0; index < Math.min(nodeOrders.length, nodes.length); index++) {
			childOrder.set(at(nodes, index), at(nodeOrders, index));
		}
	}
}

/** Return the original direct-child local-name order for a parsed XML node. */
export function smartArtChildOrder(node: XmlObject): string[] | undefined {
	return childOrder.get(node)?.slice();
}

/** Return paragraph text items in source order. */
export function orderedSmartArtTextEntries(paragraph: XmlObject): Array<[string, unknown]> {
	const keysByName = new Map<string, string>();
	for (const key of Object.keys(paragraph)) {
		if (!key.startsWith('@_')) {
			keysByName.set(localName(key), key);
		}
	}
	const itemsFor = (key: string): unknown[] => {
		const value = paragraph[key];
		return Array.isArray(value) ? value : value === undefined ? [] : [value];
	};
	const order = childOrder.get(paragraph);
	if (!order) {
		return [...keysByName.values()].flatMap((key) =>
			itemsFor(key).map((item) => [key, item] as [string, unknown]),
		);
	}
	const consumed = new Map<string, number>();
	return order.flatMap((name) => {
		const key = keysByName.get(name);
		if (!key) {
			return [];
		}
		const index = consumed.get(key) ?? 0;
		consumed.set(key, index + 1);
		const item = itemsFor(key)[index];
		return item === undefined ? [] : [[key, item]];
	});
}
