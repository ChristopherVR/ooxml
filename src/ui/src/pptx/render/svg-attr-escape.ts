/**
 * svg-attr-escape.ts: the one escaper for string-built SVG attribute values.
 *
 * A module of its own so that `silhouette-effects`, which `visual-effects`
 * imports, can escape its filter ids without an import cycle.
 *
 * @module render/svg-attr-escape
 */

/**
 * Escape a string for safe inclusion in an SVG/XML attribute value. Applied
 * to element-derived ids before they're interpolated into hand-built
 * `<filter>` markup (some bindings inject that markup via `innerHTML`/
 * `v-html`, so an unescaped id from a crafted OOXML shape id could otherwise
 * break out of the attribute).
 *
 * This is the ONE escaper for every string-concatenated SVG in `render/`.
 * `chart-sparkline`, `svg-gradient-paint` and `image-tiling` all build markup
 * that lands in `innerHTML`, and each grew its own private copy of exactly
 * this function; the copies are the failure mode, because a hardening applied
 * to one of four escapers protects a quarter of the surface.
 *
 * `String(value)` is deliberate rather than decorative: these builders are
 * fed descriptors assembled from parsed OOXML, so a field typed `string` can
 * still arrive `undefined` from a malformed deck. Coercing yields an inert
 * `"undefined"` in the attribute instead of throwing partway through building
 * a markup string, which is what one of the private copies already did and
 * what the consolidated version must keep doing.
 */
export function escapeSvgAttr(value: string): string {
	return String(value)
		.replace(/&/g, '&amp;')
		.replace(/"/g, '&quot;')
		.replace(/</g, '&lt;')
		.replace(/>/g, '&gt;');
}
