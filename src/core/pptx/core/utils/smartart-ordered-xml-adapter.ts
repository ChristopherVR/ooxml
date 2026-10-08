/**
 * The pptx boundary of the SmartArt layout interpreter: turns a
 * `fast-xml-parser` object (a layout definition's `rawXml` slot) into the
 * ordered-XML element the `diagram` walkers read (`diagram/layout/
 * smartart-choose-xml.ts`). Each object is converted once and cached; the
 * layout definition's raw slots are never mutated in place (the save path
 * merges into a fresh parse of the part), so the cache cannot go stale.
 *
 * Children keep the object tree's order (same-named siblings grouped in
 * first-appearance order), which is the order the walkers iterate anyway.
 * Element and attribute names lose their namespace prefix, as in
 * `parseOrderedXml`; `xmlns` declarations and text nodes are dropped.
 */

import type { OrderedXmlElement } from '../../../diagram/engine/ordered-xml';
import type { RawXmlView } from '../../../diagram/layout/smartart-choose-xml';
import type { XmlObject } from '../types';

const converted = new WeakMap<object, Map<string, OrderedXmlElement>>();

const localName = (key: string): string => key.slice(key.indexOf(':') + 1);

function convert(value: XmlObject, name: string): OrderedXmlElement {
	let byName = converted.get(value);
	const cached = byName?.get(name);
	if (cached) {
		return cached;
	}
	const element: OrderedXmlElement = { name, attrs: {}, children: [] };
	for (const [key, entry] of Object.entries(value)) {
		if (key.startsWith('@_')) {
			const attribute = key.slice(2);
			if (attribute !== 'xmlns' && !attribute.startsWith('xmlns:') && entry != null) {
				element.attrs[localName(attribute)] = String(entry);
			}
			continue;
		}
		if (key.startsWith('#') || key.startsWith('?')) {
			continue;
		}
		const childName = localName(key);
		for (const child of Array.isArray(entry) ? entry : [entry]) {
			element.children.push(
				child && typeof child === 'object'
					? convert(child as XmlObject, childName)
					: { name: childName, attrs: {}, children: [] },
			);
		}
	}
	if (!byName) {
		byName = new Map();
		converted.set(value, byName);
	}
	byName.set(name, element);
	return element;
}

/** The ordered-XML view of a pptx raw slot (`undefined` for a missing or non-object slot). */
export const pptxOrderedXml: RawXmlView<XmlObject> = (raw) =>
	raw && typeof raw === 'object' ? convert(raw, '') : undefined;
