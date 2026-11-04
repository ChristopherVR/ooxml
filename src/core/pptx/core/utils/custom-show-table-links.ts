import type { XmlObject } from '../types';
import { isXmlNode, xmlAttr } from './xml-access';
import { rewriteCustomShowAction, type CustomShowActionRewrite } from './custom-show-link';

type XmlValue = XmlObject[string];
interface MappedXml {
	value: XmlValue;
	changed: number;
}
const LINKS = new Set(['a:hlinkClick', 'a:hlinkMouseOver', 'a:hlinkHover']);

/**
 * Table-cell hyperlinks have no typed run fields yet. Update their preserved DrawingML,
 * copying only changed ancestors, so rich text, fields and formatting remain untouched.
 * The table frame's own action is modeled separately and must not be counted again.
 */
export function mapCustomShowTableLinks(rawXml: XmlObject, rewrite: CustomShowActionRewrite) {
	/** Rewrite actual hyperlink nodes only inside the preserved table subtree. */
	function visit(value: XmlValue, inTable: boolean, isLink = false): MappedXml {
		if (Array.isArray(value)) {
			const mapped = value.map((item) => visit(item, inTable, isLink));
			const changed = mapped.reduce((count, item) => count + item.changed, 0);
			if (!changed) return { value, changed };
			const items = mapped.flatMap((item) => (isXmlNode(item.value) ? [item.value] : []));
			return { value: items.length ? items : undefined, changed };
		}
		if (!isXmlNode(value)) return { value, changed: 0 };
		if (inTable && isLink) {
			const result = rewriteCustomShowAction(xmlAttr(value, 'action'), rewrite);
			if (result.changed)
				return {
					value: result.action === undefined ? undefined : { ...value, '@_action': result.action },
					changed: 1,
				};
		}
		let next = value;
		let changed = 0;
		for (const [key, child] of Object.entries(value)) {
			const mapped = visit(child, inTable || key === 'a:tbl', LINKS.has(key));
			if (!mapped.changed) continue;
			if (next === value) next = { ...value };
			if (mapped.value === undefined) delete next[key];
			else next[key] = mapped.value;
			changed += mapped.changed;
		}
		return { value: next, changed };
	}
	const result = visit(rawXml, false);
	return { rawXml: result.value as XmlObject, changed: result.changed };
}
