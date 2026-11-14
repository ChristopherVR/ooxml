import type { KinsokuOverride, KinsokuSettings } from '../text/kinsoku';
import { onOffElement } from './simple-types';
import { first, getW, type XmlElement } from './xml';

/**
 * The document-wide kinsoku settings of `w:settings`: `w:strictFirstAndLastChars` and the custom
 * `w:noLineBreaksBefore` / `w:noLineBreaksAfter` sets (ECMA-376 Part 1 17.15.1.58-59, 17.15.1.86).
 * Read only: settings.xml is saved as it was loaded.
 */
export function parseKinsokuSettings(root: XmlElement | null): KinsokuSettings | undefined {
	if (!root) return undefined;
	const override = (name: string): KinsokuOverride | undefined => {
		const element = first(root, name);
		const language = getW(element, 'lang');
		const chars = getW(element, 'val');
		return language && chars !== undefined ? { language, chars } : undefined;
	};
	const before = override('noLineBreaksBefore');
	const after = override('noLineBreaksAfter');
	const strict = onOffElement(first(root, 'strictFirstAndLastChars')) === true;
	if (!before && !after && !strict) return undefined;
	return {
		...(strict ? { strict: true } : {}),
		...(before ? { noLineBreaksBefore: before } : {}),
		...(after ? { noLineBreaksAfter: after } : {}),
	};
}
