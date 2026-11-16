/**
 * Framework-agnostic colour utilities for the PPTX editor.
 *
 * Colour choices resolve through the shared DrawingML resolver (`./color-transforms.ts` adapts the
 * object tree to it); this file provides the high-level colour-parsing API.
 */
import { DEFAULT_SCHEME_COLOR_MAP } from '../constants';
import type { XmlObject } from '../types';
import {
	DEFAULT_SCHEME_COLOR_CHOICE_ORDER,
	resolveObjectColorChoice,
	resolveObjectColorOpacity,
} from './color-transforms';

/**
 * Parse an OOXML colour-choice group node and resolve it to a `#RRGGBB` hex string: `a:scrgbClr`,
 * `a:srgbClr`, `a:sysClr`, `a:prstClr`, `a:schemeClr` (against the default scheme) and `a:hslClr`, with their colour transforms applied.
 *
 * @param colorNode - The XML node containing a colour-choice child element.
 * @param schemeColorOverrides - Scheme slots (lower-cased keys) that replace the default scheme.
 * @returns The resolved `#RRGGBB` hex colour string, or `undefined` if parsing fails.
 */
export function parseDrawingColorChoice(
	colorNode: XmlObject | undefined,
	schemeColorOverrides?: Readonly<Record<string, string | undefined>>,
): string | undefined {
	return resolveObjectColorChoice(
		colorNode,
		(schemeKey) => schemeColorOverrides?.[schemeKey] ?? DEFAULT_SCHEME_COLOR_MAP[schemeKey],
		undefined,
		DEFAULT_SCHEME_COLOR_CHOICE_ORDER,
	);
}

/**
 * Parse a drawing colour from an XML node, checking both direct
 * colour-choice children and an `a:solidFill` wrapper.
 *
 * This is the primary colour-parsing entry point for most OOXML
 * elements (shapes, lines, text runs) where the colour may be
 * specified either directly or inside a solid fill.
 *
 * @param colorNode - The XML node that may contain colour information.
 * @param schemeColorOverrides - Scheme slots (lower-cased keys) that replace the default scheme.
 * @returns The resolved `#RRGGBB` hex colour string, or `undefined`.
 */
export function parseDrawingColor(
	colorNode: XmlObject | undefined,
	schemeColorOverrides?: Readonly<Record<string, string | undefined>>,
): string | undefined {
	if (!colorNode) {
		return undefined;
	}
	// Try direct colour-choice first (e.g. a:srgbClr at the same level)
	const direct = parseDrawingColorChoice(colorNode, schemeColorOverrides);
	if (direct) {
		return direct;
	}
	// Fall back to checking inside an a:solidFill wrapper
	if (colorNode['a:solidFill']) {
		return parseDrawingColorChoice(colorNode['a:solidFill'] as XmlObject, schemeColorOverrides);
	}
	return undefined;
}

/**
 * The opacity (alpha) a colour-choice holder declares, read by the shared resolver from the
 * colour element {@link parseDrawingColorChoice} uses: `a:alpha` (default 1), times `a:alphaMod`,
 * plus `a:alphaOff`, each step clamped to [0, 1].
 *
 * @param colorNode - The XML node containing a colour-choice child.
 * @returns Opacity in [0, 1], or `undefined` if no alpha attributes are present.
 */
export function parseDrawingColorOpacity(colorNode: XmlObject | undefined): number | undefined {
	return resolveObjectColorOpacity(colorNode, DEFAULT_SCHEME_COLOR_CHOICE_ORDER);
}
