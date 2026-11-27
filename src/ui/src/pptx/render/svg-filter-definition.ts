/**
 * A self-contained SVG filter: its id, the CSS reference and the `<filter>` markup to place in a
 * `<defs>` (or a standalone hidden `<svg>`). Its own module so filter builders that need nothing
 * else of `visual-effects` (the bevel lighting other products' painters use) stay free of it.
 */
export interface SvgFilterDefinition {
	/** The `<filter>` element id. */
	id: string;
	/** A ready-to-use `filter: url(#id)` CSS reference. */
	cssReference: string;
	/** The `<filter>…</filter>` markup (no wrapping `<svg>`/`<defs>`). */
	filterMarkup: string;
}
