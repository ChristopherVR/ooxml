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

interface OrderCursor {
	name: string;
	nodes: XmlObject[];
	orders: string[][];
	open: string | undefined;
	depth: number;
	order: string[];
}

/** Attach source item order to parsed SmartArt paragraph objects. */
export function annotateSmartArtTextOrder(xml: string, parsed: unknown): void {
	const paragraphs = collectParsedParagraphs(parsed);
	if (paragraphs.length === 0) return;
	const items = new Map<string, XmlObject[]>([
		['r', []],
		['fld', []],
		['br', []],
	]);
	for (const paragraph of paragraphs) {
		for (const [key, value] of Object.entries(paragraph)) {
			const nodes = items.get(localName(key));
			if (!nodes) continue;
			for (const node of Array.isArray(value) ? value : [value]) {
				if (isXmlObject(node)) nodes.push(node);
			}
		}
	}
	const cursors: OrderCursor[] = [['p', paragraphs] as const, ...items]
		.filter(([, nodes]) => nodes.length > 0)
		.map(([name, nodes]) => ({ name, nodes, orders: [], open: undefined, depth: 0, order: [] }));
	// Pair source orders with the parsed nodes by occurrence, using one tag scan for every
	// collection. Empty tags parse to objects only when they have attributes.
	for (const tag of xmlTags(xml)) {
		for (const cursor of cursors) {
			if (cursor.orders.length >= cursor.nodes.length) continue;
			if (cursor.open === undefined) {
				if (tag.closing || tag.local !== cursor.name) continue;
				if (tag.selfClosing) {
					if (tag.attributes.trim()) cursor.orders.push([]);
				} else {
					cursor.open = tag.name;
					cursor.depth = 0;
					cursor.order = [];
				}
			} else if (tag.closing) {
				if (tag.name === cursor.open) {
					cursor.orders.push(cursor.order);
					cursor.open = undefined;
				} else cursor.depth = Math.max(0, cursor.depth - 1);
			} else {
				if (cursor.depth === 0) cursor.order.push(tag.local);
				if (!tag.selfClosing) cursor.depth++;
			}
		}
	}
	for (const { nodes, orders } of cursors) {
		for (const [index, order] of orders.entries()) childOrder.set(at(nodes, index), order);
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
