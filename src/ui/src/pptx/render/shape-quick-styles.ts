/**
 * Pre-defined shape quick-style presets (inspired by PowerPoint's Shape Styles
 * gallery). Framework-free pure data, drawn by the inspector's Quick Styles
 * gallery (`ribbon-galleries/quick-styles-gallery`) in every binding. Colours
 * are `#RRGGBB` with any transparency in its own field, as the model stores
 * them: the writer copies a colour into `a:srgbClr/@val` as it is.
 */
import { colorWithOpacity } from 'ooxml-core/color';
import type { PptxElement, ShapeStyle } from 'ooxml-core/pptx';
import { hasShapeProperties } from 'ooxml-core/pptx';
import { translateOr, type Translate } from 'ooxml-core/i18n';

/** A named preset shape style used in the Quick Styles gallery. */
export interface ShapeQuickStyle {
	/** English name (fallback for {@link shapeQuickStyleLabel}). */
	name: string;
	/** Dictionary key of the name (`pptx.shape.quickStylePreset.*`). */
	labelKey: string;
	style: Partial<ShapeStyle>;
}

/** The CSS a gallery swatch needs to preview a preset without rendering a shape. */
export interface QuickStyleSwatchCss {
	background: string;
	boxShadow?: string;
	border?: string;
}

/**
 * Derive the swatch CSS for a preset button.
 *
 * WHY here and not in each binding's template: the fallback order (gradient
 * beats flat colour, absent stroke means no border, shadow offsets default to
 * 2/2/4) is a presentation *rule*, and four hand-written copies of it is how
 * galleries end up looking different per framework.
 */
export function quickStyleSwatchCss(quickStyle: ShapeQuickStyle): QuickStyleSwatchCss {
	const { style } = quickStyle;
	return {
		background: style.fillGradient || style.fillColor || 'transparent',
		boxShadow: style.shadowColor
			? `${style.shadowOffsetX ?? 2}px ${style.shadowOffsetY ?? 2}px ${style.shadowBlur ?? 4}px ${colorWithOpacity(style.shadowColor, style.shadowOpacity)}`
			: undefined,
		border: style.strokeColor
			? `${style.strokeWidth ?? 1}px solid ${style.strokeColor}`
			: undefined,
	};
}

/** Serialise {@link quickStyleSwatchCss} into an inline `style` attribute value. */
export function quickStyleSwatchStyleAttr(quickStyle: ShapeQuickStyle): string {
	const css = quickStyleSwatchCss(quickStyle);
	const parts = [`background:${css.background}`];
	if (css.boxShadow) {
		parts.push(`box-shadow:${css.boxShadow}`);
	}
	if (css.border) {
		parts.push(`border:${css.border}`);
	}
	return parts.join(';');
}

/**
 * Build the element patch that applies a quick style.
 *
 * Merges over the element's existing `shapeStyle` rather than replacing it, so
 * a preset that sets no shadow leaves an author's shadow alone (matching
 * PowerPoint, where a quick style is an overlay and not a reset).
 */
export function shapeQuickStylePatch(
	el: PptxElement,
	style: Partial<ShapeStyle>,
): Partial<PptxElement> {
	const base = hasShapeProperties(el) ? (el.shapeStyle ?? {}) : {};
	return { shapeStyle: { ...base, ...style } } as Partial<PptxElement>;
}

/** A preset's name in the viewer's locale, for its swatch's `title` and `aria-label`. */
export function shapeQuickStyleLabel(quickStyle: ShapeQuickStyle, translate?: Translate): string {
	return translateOr(translate, quickStyle.labelKey, quickStyle.name);
}

export const SHAPE_QUICK_STYLES: ShapeQuickStyle[] = [
	// Row 1: Solid fills (theme-inspired palette)
	{
		name: 'Blue Fill',
		labelKey: 'pptx.shape.quickStylePreset.blueFill',
		style: { fillMode: 'solid', fillColor: '#4472C4', strokeColor: '#2F5597', strokeWidth: 1 },
	},
	{
		name: 'Orange Fill',
		labelKey: 'pptx.shape.quickStylePreset.orangeFill',
		style: { fillMode: 'solid', fillColor: '#ED7D31', strokeColor: '#C55A11', strokeWidth: 1 },
	},
	{
		name: 'Grey Fill',
		labelKey: 'pptx.shape.quickStylePreset.greyFill',
		style: { fillMode: 'solid', fillColor: '#A5A5A5', strokeColor: '#7F7F7F', strokeWidth: 1 },
	},
	{
		name: 'Gold Fill',
		labelKey: 'pptx.shape.quickStylePreset.goldFill',
		style: { fillMode: 'solid', fillColor: '#FFC000', strokeColor: '#BF9000', strokeWidth: 1 },
	},
	{
		name: 'Green Fill',
		labelKey: 'pptx.shape.quickStylePreset.greenFill',
		style: { fillMode: 'solid', fillColor: '#70AD47', strokeColor: '#548235', strokeWidth: 1 },
	},
	{
		name: 'Teal Fill',
		labelKey: 'pptx.shape.quickStylePreset.tealFill',
		style: { fillMode: 'solid', fillColor: '#5B9BD5', strokeColor: '#2E75B6', strokeWidth: 1 },
	},
	// Row 2: Outline-only styles
	{
		name: 'Blue Outline',
		labelKey: 'pptx.shape.quickStylePreset.blueOutline',
		style: { fillMode: 'none', fillColor: 'transparent', strokeColor: '#4472C4', strokeWidth: 2 },
	},
	{
		name: 'Orange Outline',
		labelKey: 'pptx.shape.quickStylePreset.orangeOutline',
		style: { fillMode: 'none', fillColor: 'transparent', strokeColor: '#ED7D31', strokeWidth: 2 },
	},
	{
		name: 'Grey Outline',
		labelKey: 'pptx.shape.quickStylePreset.greyOutline',
		style: { fillMode: 'none', fillColor: 'transparent', strokeColor: '#A5A5A5', strokeWidth: 2 },
	},
	{
		name: 'Gold Outline',
		labelKey: 'pptx.shape.quickStylePreset.goldOutline',
		style: { fillMode: 'none', fillColor: 'transparent', strokeColor: '#FFC000', strokeWidth: 2 },
	},
	{
		name: 'Green Outline',
		labelKey: 'pptx.shape.quickStylePreset.greenOutline',
		style: { fillMode: 'none', fillColor: 'transparent', strokeColor: '#70AD47', strokeWidth: 2 },
	},
	{
		name: 'Teal Outline',
		labelKey: 'pptx.shape.quickStylePreset.tealOutline',
		style: { fillMode: 'none', fillColor: 'transparent', strokeColor: '#5B9BD5', strokeWidth: 2 },
	},
	// Row 3: Fills with shadow
	{
		name: 'Blue Shadow',
		labelKey: 'pptx.shape.quickStylePreset.blueShadow',
		style: {
			fillMode: 'solid',
			fillColor: '#4472C4',
			strokeWidth: 0,
			shadowColor: '#000000',
			shadowOpacity: 0.35,
			shadowBlur: 6,
			shadowOffsetX: 2,
			shadowOffsetY: 2,
		},
	},
	{
		name: 'Orange Shadow',
		labelKey: 'pptx.shape.quickStylePreset.orangeShadow',
		style: {
			fillMode: 'solid',
			fillColor: '#ED7D31',
			strokeWidth: 0,
			shadowColor: '#000000',
			shadowOpacity: 0.35,
			shadowBlur: 6,
			shadowOffsetX: 2,
			shadowOffsetY: 2,
		},
	},
	{
		name: 'Dark Fill',
		labelKey: 'pptx.shape.quickStylePreset.darkFill',
		style: {
			fillMode: 'solid',
			fillColor: '#2F5597',
			strokeWidth: 0,
			shadowColor: '#000000',
			shadowOpacity: 0.4,
			shadowBlur: 8,
			shadowOffsetX: 3,
			shadowOffsetY: 3,
		},
	},
	{
		name: 'Dark Orange',
		labelKey: 'pptx.shape.quickStylePreset.darkOrange',
		style: {
			fillMode: 'solid',
			fillColor: '#C55A11',
			strokeWidth: 0,
			shadowColor: '#000000',
			shadowOpacity: 0.4,
			shadowBlur: 8,
			shadowOffsetX: 3,
			shadowOffsetY: 3,
		},
	},
	{
		name: 'Dark Green',
		labelKey: 'pptx.shape.quickStylePreset.darkGreen',
		style: {
			fillMode: 'solid',
			fillColor: '#548235',
			strokeWidth: 0,
			shadowColor: '#000000',
			shadowOpacity: 0.4,
			shadowBlur: 8,
			shadowOffsetX: 3,
			shadowOffsetY: 3,
		},
	},
	{
		name: 'White Fill',
		labelKey: 'pptx.shape.quickStylePreset.whiteFill',
		style: {
			fillMode: 'solid',
			fillColor: '#FFFFFF',
			strokeColor: '#D9D9D9',
			strokeWidth: 1,
			shadowColor: '#000000',
			shadowOpacity: 0.2,
			shadowBlur: 4,
			shadowOffsetX: 1,
			shadowOffsetY: 1,
		},
	},
];
