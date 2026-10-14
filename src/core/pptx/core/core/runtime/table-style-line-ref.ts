/**
 * Resolve a table style's `a:lnRef` border (CT_StyleMatrixReference, the
 * choice sibling of `a:ln` in `a:tcBdr` sides) against the theme's line
 * styles.
 *
 * "Themed Style 1" and "Themed Style 2" draw their gridlines through
 * `<a:lnRef idx="1"><a:schemeClr val="accentN"/></a:lnRef>` rather than a
 * concrete `a:ln`, so the line's width and dash live in the theme's
 * `a:fmtScheme/a:lnStyleLst` and its colour is the reference's own colour
 * substituted for the style's `phClr`. A pure function over plain data, like
 * `resolveTableBackgroundFillRef` for `a:tblBg/a:fillRef`.
 *
 * @module table-style-line-ref
 */
import { applyDrawingColorTransforms } from '../../color';
import type {
	ParsedTableFillRef,
	ParsedTableStyleBorder,
	PptxThemeFormatScheme,
	XmlObject,
} from '../../types';

/** Whether a raw XML node has the child `key` at all (`<a:noFill/>` parses as ''). */
function hasChild(rawNode: unknown, key: string): boolean {
	return (
		!!rawNode &&
		typeof rawNode === 'object' &&
		Object.prototype.hasOwnProperty.call(rawNode, key) &&
		(rawNode as XmlObject)[key] !== undefined
	);
}

/** The `a:schemeClr val="phClr"` (with its transforms) of a format-scheme line's solid fill. */
function linePlaceholderColor(rawNode: unknown): XmlObject | undefined {
	if (!rawNode || typeof rawNode !== 'object') {
		return undefined;
	}
	const solidFill = (rawNode as XmlObject)['a:solidFill'] as XmlObject | undefined;
	const schemeClr = solidFill?.['a:schemeClr'] as XmlObject | undefined;
	return schemeClr?.['@_val'] === 'phClr' ? schemeClr : undefined;
}

/**
 * The concrete border an `a:lnRef` stands for, or `undefined` when the theme
 * has no line style at that index. `idx` 0 means no line (ECMA-376
 * 20.1.4.2.19). `colorMap` maps scheme keys (`accent1`, ...) to hex.
 */
export function resolveTableStyleLineRef(
	lnRef: ParsedTableFillRef,
	formatScheme: PptxThemeFormatScheme | undefined,
	colorMap: Readonly<Record<string, string>> | undefined,
): ParsedTableStyleBorder | undefined {
	if (lnRef.idx === 0) {
		return { noFill: true };
	}
	const line = formatScheme?.lineStyles[lnRef.idx - 1];
	if (!line) {
		return undefined;
	}
	// A theme line that says "no line" stays no line: without this the border
	// came back with only a width, and the border renderer defaults a colourless
	// border to black. Key presence, not truthiness: the parser reads
	// `<a:noFill/>` as ''.
	if (hasChild(line.rawNode, 'a:noFill')) {
		return { noFill: true };
	}
	const border: ParsedTableStyleBorder = {};
	if (line.width !== undefined) {
		border.width = line.width;
	}
	if (line.dash && line.dash !== 'solid') {
		border.dash = line.dash;
	}
	// The reference's colour stands in for the line's `phClr` only; a line with
	// a colour of its own keeps it.
	const placeholder = linePlaceholderColor(line.rawNode);
	const base = placeholder ? refColor(lnRef, colorMap) : undefined;
	if (base && placeholder) {
		border.color = applyDrawingColorTransforms(base, placeholder);
	} else if (line.color) {
		border.color = line.color;
	}
	return border;
}

/** The reference's own colour (`a:schemeClr` with its tint/shade/alpha) as hex. */
function refColor(
	lnRef: ParsedTableFillRef,
	colorMap: Readonly<Record<string, string>> | undefined,
): string | undefined {
	const color = lnRef.color;
	const base = color?.schemeColor ? colorMap?.[color.schemeColor] : color?.color;
	if (!base || !color) {
		return base;
	}
	const transforms: XmlObject = {};
	for (const key of ['tint', 'shade', 'alpha'] as const) {
		if (color[key] !== undefined) {
			transforms[`a:${key}`] = { '@_val': String(color[key]) };
		}
	}
	return Object.keys(transforms).length > 0 ? applyDrawingColorTransforms(base, transforms) : base;
}
