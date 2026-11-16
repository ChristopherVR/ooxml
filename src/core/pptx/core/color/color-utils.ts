/**
 * Framework-agnostic colour utilities for the PPTX editor.
 *
 * Colour choices resolve through the shared DrawingML resolver (`./color-transforms.ts` adapts the
 * object tree to it); this file provides the high-level colour-parsing API.
 */
import {
	clampUnitInterval,
	parseDrawingPercent,
	parseDrawingFraction,
} from '../../../color/color-primitives';
import { DEFAULT_SCHEME_COLOR_MAP } from '../constants';
import type { XmlObject } from '../types';
import { DEFAULT_SCHEME_COLOR_CHOICE_ORDER, resolveObjectColorChoice } from './color-transforms';

/**
 * Parse an OOXML colour-choice group node and resolve it to a `#RRGGBB` hex string: `a:scrgbClr`,
 * `a:srgbClr`, `a:sysClr`, `a:prstClr`, `a:schemeClr` (against the default scheme) and `a:hslClr`, with their colour transforms applied.
 *
 * @param colorNode - The XML node containing a colour-choice child element.
 * @returns The resolved `#RRGGBB` hex colour string, or `undefined` if parsing fails.
 */
export function parseDrawingColorChoice(colorNode: XmlObject | undefined): string | undefined {
	return resolveObjectColorChoice(
		colorNode,
		(schemeKey) => DEFAULT_SCHEME_COLOR_MAP[schemeKey],
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
 * @returns The resolved `#RRGGBB` hex colour string, or `undefined`.
 */
export function parseDrawingColor(colorNode: XmlObject | undefined): string | undefined {
	if (!colorNode) {
		return undefined;
	}
	// Try direct colour-choice first (e.g. a:srgbClr at the same level)
	const direct = parseDrawingColorChoice(colorNode);
	if (direct) {
		return direct;
	}
	// Fall back to checking inside an a:solidFill wrapper
	if (colorNode['a:solidFill']) {
		return parseDrawingColorChoice(colorNode['a:solidFill'] as XmlObject);
	}
	return undefined;
}

/**
 * Extract the opacity (alpha) value from an OOXML colour node.
 *
 * Checks all six colour-choice types for `a:alpha`, `a:alphaMod`,
 * and `a:alphaOff` child elements, and combines them:
 *   - `alpha` sets the base opacity (default 1.0 if absent)
 *   - `alphaMod` multiplies the opacity
 *   - `alphaOff` adds an offset to the opacity
 *
 * @param colorNode - The XML node containing a colour-choice child.
 * @returns Opacity in [0, 1], or `undefined` if no alpha attributes are present.
 */
export function parseDrawingColorOpacity(colorNode: XmlObject | undefined): number | undefined {
	if (!colorNode) {
		return undefined;
	}

	// Find whichever colour-choice type is present
	const colorChoice =
		(colorNode['a:scrgbClr'] as XmlObject | undefined) ||
		(colorNode['a:srgbClr'] as XmlObject | undefined) ||
		(colorNode['a:schemeClr'] as XmlObject | undefined) ||
		(colorNode['a:hslClr'] as XmlObject | undefined) ||
		(colorNode['a:prstClr'] as XmlObject | undefined) ||
		(colorNode['a:sysClr'] as XmlObject | undefined);
	if (!colorChoice) {
		return undefined;
	}

	// Parse the three alpha-related transform values. `alphaOff` is an additive
	// offset that may be negative (e.g. `-25%`), so it must be parsed unclamped;
	// clamping it to [0, 1] first would silently discard negative offsets. The
	// combined opacity is clamped once at the end.
	const alpha = parseDrawingPercent((colorChoice['a:alpha'] as XmlObject | undefined)?.['@_val']);
	const alphaMod = parseDrawingPercent(
		(colorChoice['a:alphaMod'] as XmlObject | undefined)?.['@_val'],
	);
	const alphaOff = parseDrawingFraction(
		(colorChoice['a:alphaOff'] as XmlObject | undefined)?.['@_val'],
	);
	if (alpha === undefined && alphaMod === undefined && alphaOff === undefined) {
		return undefined;
	}

	// Combine: start with absolute alpha (or 1.0), multiply by mod, add offset
	let opacity = alpha ?? 1;
	if (alphaMod !== undefined) {
		opacity *= alphaMod;
	}
	if (alphaOff !== undefined) {
		opacity += alphaOff;
	}
	return clampUnitInterval(opacity);
}
