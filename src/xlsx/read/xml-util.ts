import { NS, children, elements, first, type XmlElement } from '../../xml/index.js';
import { XMLSerializer } from '@xmldom/xmldom';

type XmldomNode = Parameters<InstanceType<typeof XMLSerializer>['serializeToString']>[0];

/** Direct SpreadsheetML children named `local`. */
export const xChildren = (parent: ParentNode, local: string): XmlElement[] =>
	children(parent, local, NS.x);

/** The first SpreadsheetML child named `local`. */
export const xFirst = (
	parent: ParentNode | null | undefined,
	local: string,
): XmlElement | undefined => first(parent, local, NS.x);

/** An un-namespaced attribute (SpreadsheetML attributes carry no prefix). */
export const att = (element: XmlElement | null | undefined, name: string): string | undefined => {
	if (!element || !element.hasAttribute(name)) return undefined;
	return element.getAttribute(name) ?? undefined;
};

/** `1`/`true` -> true, `0`/`false` -> false, absent -> `fallback`. */
export function boolAttr(element: XmlElement | null | undefined, name: string): boolean | undefined;
export function boolAttr(
	element: XmlElement | null | undefined,
	name: string,
	fallback: boolean,
): boolean;
export function boolAttr(
	element: XmlElement | null | undefined,
	name: string,
	fallback?: boolean,
): boolean | undefined {
	const value = att(element, name);
	if (value === undefined) return fallback;
	return value === '1' || value.toLowerCase() === 'true';
}

/** A numeric attribute, or `undefined` when absent or not a finite number. */
export function numAttr(element: XmlElement | null | undefined, name: string): number | undefined {
	const value = att(element, name);
	if (value === undefined || value.trim() === '') return undefined;
	const n = Number(value);
	return Number.isFinite(n) ? n : undefined;
}

/** The `val` attribute of a child element (`<sz val="11"/>`). */
export const childVal = (parent: XmlElement, local: string): string | undefined =>
	att(xFirst(parent, local), 'val');

/** Text of an element, decoding SpreadsheetML `_xHHHH_` escapes. */
export const xText = (element: XmlElement | null | undefined): string =>
	decodeEscapes(element?.textContent ?? '');

/** Decodes `_xHHHH_` escapes (`_x000D_` carriage return, `_x005F_` a literal underscore). */
export function decodeEscapes(text: string): string {
	if (!text.includes('_x')) return text;
	return text.replace(/_x([0-9A-Fa-f]{4})_/g, (_, hex: string) =>
		String.fromCharCode(Number.parseInt(hex, 16)),
	);
}

const serializer = new XMLSerializer();

/** Serializes an element on its own; xmldom adds the namespace declarations it needs. */
export const outerXml = (element: XmlElement): string =>
	serializer.serializeToString(element as unknown as XmldomNode);

/**
 * Serializes an `mc:AlternateContent` (or any element) so that every prefix named in a
 * `Requires` attribute is declared on the element itself, keeping the fragment self-contained.
 */
export function selfContainedXml(element: XmlElement): string {
	const clone = element.cloneNode(true) as XmlElement;
	const visit = (node: XmlElement) => {
		const requires = node.getAttribute('Requires');
		if (requires) {
			for (const prefix of requires.split(/\s+/).filter(Boolean)) {
				const uri = element.lookupNamespaceURI(prefix) ?? node.lookupNamespaceURI(prefix);
				if (uri && !clone.hasAttribute(`xmlns:${prefix}`))
					clone.setAttribute(`xmlns:${prefix}`, uri);
			}
		}
		for (const child of elements(node)) visit(child);
	};
	visit(element);
	return outerXml(clone);
}

/** All descendant elements named `local` in `ns` (document order). */
export function descendants(root: XmlElement, local: string, ns: string): XmlElement[] {
	return Array.from(root.getElementsByTagNameNS(ns, local));
}
