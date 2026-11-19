/**
 * css-escape.ts: escape a value for use inside an attribute selector.
 *
 * Element ids and similar keys can come from a file or a collaboration peer,
 * and the escaped text lands both in `querySelector` calls and in `<style>`
 * text. Hex escapes serve both: the CSS parser reads back the original
 * character, the selector cannot be closed early, and the text holds nothing
 * an HTML parser would act on. This is the one escaper for that job.
 *
 * @module render/css-escape
 */

const CSS_STRING_UNSAFE = /["\\<>&'\u0000-\u001f\u007f]/gu;

/** Escape `value` for a double-quoted CSS string, as `\3c ` style hex escapes. */
export function escapeCssAttributeValue(value: string): string {
	return value.replace(CSS_STRING_UNSAFE, (char) => `\\${char.charCodeAt(0).toString(16)} `);
}

/** `[data-element-id="<id>"]`, with the id escaped for the selector. */
export function elementIdSelector(id: string): string {
	return `[data-element-id="${escapeCssAttributeValue(id)}"]`;
}
