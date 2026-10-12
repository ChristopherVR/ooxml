/**
 * inspector-helpers.ts: Pure (no framework) helpers for the inspector panel.
 *
 * Readers extract display values from a PptxElement (with sensible defaults),
 * and patch builders produce shallow-merge-ready Partial<PptxElement> objects
 * for the binding's element-update path. Shared across React, Vue, and Angular.
 */

import type {
	PlaceholderDefaults,
	PptxElement,
	PptxThemeColorRef,
	TextStyle,
} from 'ooxml-core/pptx';
import { hasShapeProperties, hasTextProperties } from 'ooxml-core/pptx';

import { mergeTextStyleUpdate } from './font-family-update';
import { textStyleFlagsOf } from './selection-format-state';
import { textFontSizePxToPt } from './text-format-presets';
import { paragraphOverridePatch, withExclusiveLineSpacing } from './text-style-scope';
import { updateTextSegmentStyle } from './update-text-segment-style';

const DEFAULT_FILL = '#ffffff';
const DEFAULT_STROKE = '#000000';
const DEFAULT_TEXT_COLOR = '#000000';
/**
 * PowerPoint's own presentation-level default text style (`p:defaultTextStyle
 * /a:lvl1pPr/a:defRPr@sz`) is 1800 (18pt), so a shape with no explicit font
 * size renders at 18pt in real PowerPoint, not 12pt. Kept as the last-resort
 * fallback below the deck's actual default when {@link fontSizeOf} is given one.
 */
export const DEFAULT_FONT_SIZE = 18;

/**
 * Returns the fill colour of the element's shapeStyle, or a white default.
 * Only meaningful for elements that pass `hasShapeProperties`.
 */
export function fillColorOf(el: PptxElement): string {
	if (hasShapeProperties(el)) {
		return el.shapeStyle?.fillColor ?? DEFAULT_FILL;
	}
	return DEFAULT_FILL;
}

/**
 * Returns the stroke colour of the element's shapeStyle, or a black default.
 * Only meaningful for elements that pass `hasShapeProperties`.
 */
export function strokeColorOf(el: PptxElement): string {
	if (hasShapeProperties(el)) {
		return el.shapeStyle?.strokeColor ?? DEFAULT_STROKE;
	}
	return DEFAULT_STROKE;
}

/**
 * Returns the text colour from the element's textStyle, or a black default.
 * Only meaningful for elements that pass `hasTextProperties`.
 */
export function textColorOf(el: PptxElement): string {
	if (hasTextProperties(el)) {
		return el.textStyle?.color ?? DEFAULT_TEXT_COLOR;
	}
	return DEFAULT_TEXT_COLOR;
}

/**
 * Returns the font size (in points) from the element's textStyle.
 *
 * Falls back, in order, to: the deck's presentation-level default text style
 * (`p:defaultTextStyle`, passed as `presentationDefault` when the caller has
 * one), then `DEFAULT_FONT_SIZE` (18pt, matching PowerPoint's own
 * `a:lvl1pPr/a:defRPr@sz="1800"` when neither is available).
 */
export function fontSizeOf(el: PptxElement, presentationDefault?: PlaceholderDefaults): number {
	if (hasTextProperties(el) && el.textStyle?.fontSize !== undefined) {
		return textFontSizePxToPt(el.textStyle.fontSize);
	}
	const deckDefault = presentationDefault?.levelStyles?.[0]?.fontSize;
	return deckDefault === undefined ? DEFAULT_FONT_SIZE : textFontSizePxToPt(deckDefault);
}

/**
 * Returns whether the element's text is bold: every content run effectively
 * bold (run style, else the body style), false when absent. Derived from the
 * runs, not the body alone, so a box whose runs are all bold at run level
 * reads as bold and a mixed box does not (see `selection-format-state`).
 */
export function isBold(el: PptxElement): boolean {
	return textStyleFlagsOf(el).bold === 'on';
}

/** Returns whether the element's text is italic (see {@link isBold} for the run rule). */
export function isItalic(el: PptxElement): boolean {
	return textStyleFlagsOf(el).italic === 'on';
}

/** Returns whether the element's text has underline (see {@link isBold} for the run rule). */
export function isUnderline(el: PptxElement): boolean {
	return textStyleFlagsOf(el).underline === 'on';
}

/**
 * Changes to apply to shapeStyle: a subset of ShapeStyle fields editable
 * from the inspector panel.
 */
export interface ShapeStyleChanges {
	fillColor?: string;
	/**
	 * Set alongside `fillColor` when the colour came from a theme swatch (wins
	 * on save); pass `undefined` explicitly when clearing a previously-stored
	 * ref (a custom hex or recent-colour pick has no theme identity).
	 */
	fillColorRef?: PptxThemeColorRef;
	strokeColor?: string;
	/** Same contract as {@link fillColorRef}, for the stroke colour. */
	strokeColorRef?: PptxThemeColorRef;
}

/**
 * Builds a Partial<PptxElement> patch that merges the given changes into the
 * element's existing shapeStyle without dropping any other shapeStyle fields.
 */
export function shapeStylePatch(el: PptxElement, changes: ShapeStyleChanges): Partial<PptxElement> {
	const base = hasShapeProperties(el) ? (el.shapeStyle ?? {}) : {};
	return {
		shapeStyle: {
			...base,
			...changes,
		},
	} as Partial<PptxElement>;
}

/**
 * Changes to apply to textStyle: a subset of TextStyle fields editable
 * from the inspector panel.
 */
export interface TextStyleChanges {
	color?: string;
	/** Same contract as {@link ShapeStyleChanges.fillColorRef}, for the text colour. */
	colorRef?: PptxThemeColorRef;
	fontSize?: number;
	bold?: boolean;
	italic?: boolean;
	underline?: boolean;
}

/**
 * Builds a Partial<PptxElement> patch that merges the given changes into the
 * element's existing textStyle without dropping any other textStyle fields.
 */
export function textStylePatch(el: PptxElement, changes: Partial<TextStyle>): Partial<PptxElement> {
	const base = hasTextProperties(el) ? (el.textStyle ?? {}) : {};
	const exclusive = withExclusiveLineSpacing(changes);
	const paragraphs = paragraphOverridePatch(el, exclusive);
	const segments = paragraphs.textSegments ?? (hasTextProperties(el) ? el.textSegments : undefined);
	return {
		...paragraphs,
		...(segments?.some((segment) => segment.paragraphInsertionStyle)
			? {
					textSegments: segments.map((segment) =>
						updateTextSegmentStyle(segment, exclusive, { updateBodyStyle: false }),
					),
				}
			: {}),
		textStyle: mergeTextStyleUpdate(base, exclusive),
	} as Partial<PptxElement>;
}

/** Apply `updates` to the element's body style and to every one of its runs. */
function textRunsPatch(el: PptxElement, updates: Partial<TextStyle>): Partial<PptxElement> {
	const patch = textStylePatch(el, updates);
	if (!hasTextProperties(el) || !el.textSegments) {
		return patch;
	}
	return {
		...patch,
		textSegments: el.textSegments.map((segment) => updateTextSegmentStyle(segment, updates)),
	} as Partial<PptxElement>;
}

/** Apply an ordinary-text model pixel size to the element and all of its runs. */
export function textFontSizePatch(el: PptxElement, fontSize: number): Partial<PptxElement> {
	return textRunsPatch(el, { fontSize });
}

/**
 * Apply a font family to the element and all of its runs, the way
 * {@link textFontSizePatch} applies a size. Setting it on the body alone left
 * every run that names its own `a:latin` in the old face, so a whole-box pick
 * from the ribbon changed nothing; the runs also carry the same-face `a:ea` /
 * `a:cs` slots along (see `font-family-update.ts`).
 */
export function textFontFamilyPatch(el: PptxElement, fontFamily: string): Partial<PptxElement> {
	return textRunsPatch(el, { fontFamily });
}
