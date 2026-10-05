import type { XmlElement } from '../../xml/index.js';
import { attrs, escapeAttr, type AttrValue } from './xml-out.js';

/**
 * The attributes of a regenerated element: every attribute of the source element (unknown ones
 * and prefixed ones such as `x14ac:dyDescent` included, with the prefix declarations the
 * fragment carries) in source order, then the modelled `values`. A key in `values` is owned by
 * the model: `undefined` removes it, anything else replaces the source value.
 */
export function mergedAttrs(
	source: XmlElement | undefined,
	values: Record<string, AttrValue>,
): string {
	let out = '';
	const owned = new Set(Object.keys(values));
	for (const attribute of source ? Array.from(source.attributes) : []) {
		if (attribute.name === 'xmlns' || owned.has(attribute.name)) continue;
		out += ` ${attribute.name}="${escapeAttr(attribute.value)}"`;
	}
	return out + attrs(values);
}
