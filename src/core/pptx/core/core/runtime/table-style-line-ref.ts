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
import type { DrawingColor } from '../../../../drawingml/types';
import {
	applyColorTransformListTo,
	applyDrawingColorTransforms,
} from '../../color/color-transforms';
import { OBJECT_DRAWING_VIEW } from '../../color/object-drawing-view';
import { parseSolidFillStyle } from './table-style-border-parse';
import type {
	ParsedTableFillRef,
	ParsedTableStyleBorder,
	ParsedTableStyleFill,
	PptxThemeFormatScheme,
	XmlObject,
} from '../../types';

const asNode = (rawNode: unknown): XmlObject | undefined =>
	rawNode && typeof rawNode === 'object' ? (rawNode as XmlObject) : undefined;

/** Whether a theme line says "no line" (`<a:noFill/>`, read by the shared node view). */
function isNoLine(rawNode: unknown): boolean {
	const node = asNode(rawNode);
	return !!node && OBJECT_DRAWING_VIEW.elements(node).some((child) => child.name === 'noFill');
}

/** The colour of a format-scheme line's solid fill, read by the shared node view. */
function lineSolidColor(rawNode: unknown): DrawingColor | undefined {
	const node = asNode(rawNode);
	const solidFill = node && OBJECT_DRAWING_VIEW.child(node, 'solidFill');
	return solidFill ? OBJECT_DRAWING_VIEW.color(solidFill) : undefined;
}

/** The `phClr` colour (with its transforms) of a format-scheme line's solid fill. */
function linePlaceholderColor(rawNode: unknown): DrawingColor | undefined {
	const color = lineSolidColor(rawNode);
	return color?.kind === 'scheme' && color.value === 'phClr' ? color : undefined;
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
	if (isNoLine(line.rawNode)) {
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
		border.color = applyColorTransformListTo(base, placeholder.transforms);
	} else if (line.color) {
		border.color = line.color;
	}
	return border;
}

/** `a:tint`/`a:shade` compose by multiplying in linear light, so two of a kind fold into one. */
function foldPercent(a: number | undefined, b: number | undefined): number | undefined {
	if (a === undefined || b === undefined) {
		return a ?? b;
	}
	return Math.round((a * b) / 100_000);
}

/** A reference colour as a border colour: `fill` when theme-linked, `color` for sRGB. */
function refBorderColor(
	ref: ParsedTableStyleFill | undefined,
): Pick<ParsedTableStyleBorder, 'fill' | 'color'> | undefined {
	if (ref?.schemeColor) {
		return { fill: ref };
	}
	// The border writer emits sRGB from `color` only (without transforms).
	return ref?.color ? { color: ref.color } : undefined;
}

/**
 * The colour an `a:lnRef` side draws in, as a theme-linked border colour
 * (`fill` for a scheme colour, `color` for sRGB), for an edit that turns the
 * reference into a concrete line while leaving its colour alone. A theme line
 * with a colour of its own keeps that colour; a `phClr` line takes the
 * reference's colour with the line's tint, shade and alpha folded in. Other
 * transforms (`satMod`, `lumMod`, ...) have no place in the table-style fill
 * model and are dropped, as when any table-style colour is parsed. Without
 * the scheme, or a solid colour on the theme line, it is the reference's
 * colour as is.
 */
export function resolveTableStyleLineRefColor(
	lnRef: ParsedTableFillRef,
	formatScheme: PptxThemeFormatScheme | undefined,
): Pick<ParsedTableStyleBorder, 'fill' | 'color'> | undefined {
	const ref = lnRef.color;
	const line = lnRef.idx > 0 ? formatScheme?.lineStyles[lnRef.idx - 1] : undefined;
	const node = asNode(line?.rawNode);
	const solidFill = node && OBJECT_DRAWING_VIEW.child(node, 'solidFill');
	const own = parseSolidFillStyle(solidFill);
	if (!own) {
		const color = solidFill ? OBJECT_DRAWING_VIEW.color(solidFill) : undefined;
		const hex = color?.kind === 'srgb' ? color.value : '';
		return hex ? { color: hex.startsWith('#') ? hex : `#${hex}` } : refBorderColor(ref);
	}
	if (own.schemeColor !== 'phClr') {
		return { fill: own };
	}
	if (!ref?.schemeColor) {
		return refBorderColor(ref);
	}
	const fill: ParsedTableStyleFill = { schemeColor: ref.schemeColor };
	const tint = foldPercent(ref.tint, own.tint);
	const shade = foldPercent(ref.shade, own.shade);
	const alpha = own.alpha ?? ref.alpha;
	if (tint !== undefined) {
		fill.tint = tint;
	}
	if (shade !== undefined) {
		fill.shade = shade;
	}
	if (alpha !== undefined) {
		fill.alpha = alpha;
	}
	return { fill };
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
