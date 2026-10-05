// Small DOM helpers for the diagram parsers, over the shared `xml` area.
import { NS, children, elements, first, isElement, type XmlElement } from '../xml/index.js';
import type { AttributeReader } from './types.js';

export { children, elements, first, isElement, NS };
export type { XmlElement };

/** Reads an attribute by local name regardless of its prefix (`r:dm` and `dm` are the same here). */
export function attributeReader(element: XmlElement | null | undefined): AttributeReader {
	return (name) => {
		if (!element) return undefined;
		const direct = element.getAttribute(name);
		if (direct !== null) return direct;
		const attributes = element.attributes;
		for (let index = 0; index < attributes.length; index++) {
			const candidate = attributes.item(index);
			if (candidate && candidate.localName === name) return candidate.value;
		}
		return undefined;
	};
}

/** A trimmed, non-empty attribute value. */
export function stringAttribute(
	element: XmlElement | null | undefined,
	name: string,
): string | undefined {
	const value = attributeReader(element)(name)?.trim();
	return value ? value : undefined;
}

/** An integer attribute (base 10); `undefined` when absent or not a number. */
export function integerAttribute(
	element: XmlElement | null | undefined,
	name: string,
): number | undefined {
	const parsed = Number.parseInt(attributeReader(element)(name) ?? '', 10);
	return Number.isFinite(parsed) ? parsed : undefined;
}

/** `1`, `true` -> true; `0`, `false` -> false; otherwise `undefined`. */
export function booleanAttribute(
	element: XmlElement | null | undefined,
	name: string,
): boolean | undefined {
	const value = attributeReader(element)(name)?.trim().toLowerCase();
	if (value === '1' || value === 'true') return true;
	if (value === '0' || value === 'false') return false;
	return undefined;
}

/** All descendant elements with a local name, in document order, in any namespace. */
export function descendants(root: XmlElement, local: string): XmlElement[] {
	const result: XmlElement[] = [];
	const visit = (node: XmlElement) => {
		for (const child of Array.from(node.childNodes)) {
			if (!isElement(child)) continue;
			if (child.localName === local) result.push(child);
			visit(child);
		}
	};
	visit(root);
	return result;
}
