import type { XmlObject } from '../types';

/**
 * Whether `node` has a child named `name`, by KEY existence rather than
 * `getChildByLocalName`'s "resolves to a non-array object" check. A childless,
 * attribute-less element (`<c:majorGridlines/>`, `<a:noFill/>`) is a legal,
 * common way to mark a boolean flag present, but fast-xml-parser renders it
 * as an empty STRING, which `getChildByLocalName` cannot distinguish from
 * "absent".
 */
export function hasLocalName(node: XmlObject, name: string): boolean {
	if (Object.hasOwn(node, name)) {
		return true;
	}
	const suffix = `:${name}`;
	return Object.keys(node).some((key) => key.endsWith(suffix));
}
