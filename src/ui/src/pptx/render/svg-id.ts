const SVG_ID_UNSAFE = /[^A-Za-z0-9.-]/gu;

/**
 * `value` as text that is safe inside an SVG id and every `url(#id)` or `href="#id"` reference.
 *
 * Browsers percent-decode a reference's fragment but not the id it looks for, so an element id
 * holding `%` (as one built from an encoded part name does) would miss its own def. Each other
 * character becomes `_<hex>_`, `_` included, so two different values never share a token.
 */
export function svgIdToken(value: string): string {
	return value.replace(SVG_ID_UNSAFE, (char) => `_${char.codePointAt(0)!.toString(16)}_`);
}
