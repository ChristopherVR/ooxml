/**
 * OOXML drawing-colour resolution for the bindings: colour-choice parsing (srgb/scrgb/sys/scheme/
 * hsl/preset), the 26 colour transforms, scheme overrides and alpha. Every reader here delegates
 * to the core object-tree adapter of the shared DrawingML resolver, so the bindings resolve a
 * colour exactly as the runtime does (docs/agnostic-core-plan.md, step 3).
 */
import {
	XmlObject,
	applyDrawingColorTransforms as applyFullColorTransforms,
	parseDrawingColor as parseCoreDrawingColor,
	parseDrawingColorChoice as parseCoreDrawingColorChoice,
	parseDrawingColorOpacity as parseCoreDrawingColorOpacity,
} from 'ooxml-core/pptx';

export const DEFAULT_SCHEME_COLOR_MAP: Record<string, string> = {
	dk1: '#000000',
	lt1: '#FFFFFF',
	dk2: '#1F497D',
	lt2: '#EEECE1',
	accent1: '#4472C4',
	accent2: '#ED7D31',
	accent3: '#A5A5A5',
	accent4: '#FFC000',
	accent5: '#5B9BD5',
	accent6: '#70AD47',
	hlink: '#0563C1',
	folHlink: '#954F72',
	tx1: '#000000',
	tx2: '#44546A',
	bg1: '#FFFFFF',
	bg2: '#E7E6E6',
	phclr: '#4472C4',
};

/**
 * Apply all 26 OOXML drawing colour transforms to a base colour.
 * Delegates to the comprehensive implementation in the editor core.
 */
export function applyDrawingColorTransforms(baseColor: string, colorNode: XmlObject): string {
	return applyFullColorTransforms(baseColor, colorNode);
}

/**
 * Resolve a colour-choice holder to `#RRGGBB`. `schemeColorOverrides` (lower-cased scheme keys)
 * replace the default scheme slots.
 */
export function parseDrawingColorChoice(
	colorNode: XmlObject | undefined,
	schemeColorOverrides?: Record<string, string | undefined>,
): string | undefined {
	return parseCoreDrawingColorChoice(colorNode, schemeColorOverrides);
}

/** {@link parseDrawingColorChoice} on the node itself, then on its `a:solidFill`. */
export function parseDrawingColor(
	colorNode: XmlObject | undefined,
	schemeColorOverrides?: Record<string, string | undefined>,
): string | undefined {
	return parseCoreDrawingColor(colorNode, schemeColorOverrides);
}

/** The opacity the holder's colour declares (the shared resolver's alpha), or `undefined`. */
export function parseDrawingColorOpacity(colorNode: XmlObject | undefined): number | undefined {
	return parseCoreDrawingColorOpacity(colorNode);
}
