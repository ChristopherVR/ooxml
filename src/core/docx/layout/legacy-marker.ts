/**
 * Where the text after a legacy (`w:legacy`) list marker starts, measured the way the
 * marker's origin is: the text begins `indentPx` after the start of the number, and at least
 * `spacePx` after its end (ECMA-376 17.9.10, 17.9.11). Print Layout, the PDF export and the
 * editing view share this rule.
 */
export function legacyTextStartPx(
	numberStartPx: number,
	numberEndPx: number,
	legacy: { indentPx: number; spacePx: number },
): number {
	return Math.max(numberStartPx + legacy.indentPx, numberEndPx + legacy.spacePx);
}
