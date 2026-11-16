/**
 * The shape Fill options of the inspector (Format Shape > Fill & Line > Fill):
 * the fill kind (no fill, solid, gradient, picture, pattern, slide background)
 * and the plain fields each kind shows (transparency, gradient type and
 * angle, picture tiling), as field descriptors every binding draws with its
 * one generic renderer.
 *
 * Like the Line descriptors these read a `ShapeStyle` and write a
 * `Partial<ShapeStyle>`. The compound controls (the colour pickers with their
 * theme grids, the gradient stop list, the pattern swatches and the picture
 * source) are not fields: `shape-fill-actions.ts` resolves what they show and
 * turns their events into the same style patches.
 *
 * @module render/inspector/shape-fill-fields
 */
import type { ShapeStyle } from 'ooxml-core/pptx';
import { normalizeHexColor } from 'ooxml-core/color';

import { DEFAULT_FILL_COLOR } from '../../constants';
import type { InspectorFieldOption, InspectorSectionDescriptor } from './inspector-field-model';
import { shapeFillStopsOf } from './shape-fill-stops';

/** The fill kinds of PowerPoint's Fill pane, in its order. */
export type ShapeFillMode = 'none' | 'solid' | 'gradient' | 'image' | 'pattern' | 'background';

/** What the Fill descriptors read and write. */
export type ShapeFillTarget = ShapeStyle | undefined;
export type ShapeFillSection = InspectorSectionDescriptor<ShapeFillTarget, Partial<ShapeStyle>>;

/** The preset a shape switching into a pattern fill starts with. */
export const DEFAULT_PATTERN_PRESET = 'pct20';
/** The second colour of a new pattern fill. */
export const DEFAULT_PATTERN_BACKGROUND = '#ffffff';

/** The fill kinds a binding's fill select offers. */
export const SHAPE_FILL_MODE_OPTIONS: readonly InspectorFieldOption[] = [
	{ value: 'none', labelKey: 'pptx.fill.none' },
	{ value: 'solid', labelKey: 'pptx.fill.solid' },
	{ value: 'gradient', labelKey: 'pptx.fill.gradient' },
	{ value: 'image', labelKey: 'pptx.fill.picture' },
	{ value: 'pattern', labelKey: 'pptx.table.fillPattern' },
	{ value: 'background', labelKey: 'pptx.fill.slideBackground' },
];

const GRADIENT_TYPE_OPTIONS: readonly InspectorFieldOption[] = [
	{ value: 'linear', labelKey: 'pptx.gradient.linear' },
	{ value: 'radial', labelKey: 'pptx.gradient.radial' },
];

const IMAGE_MODE_OPTIONS: readonly InspectorFieldOption[] = [
	{ value: 'stretch', labelKey: 'pptx.image.stretch' },
	{ value: 'tile', labelKey: 'pptx.image.tile' },
];

/** Whether `color` is a colour a picker can show (not `transparent` or missing). */
export function isUsableFillColor(color: string | undefined): color is string {
	return (
		typeof color === 'string' && /^#?[0-9a-f]{3}(?:[0-9a-f]{3})?(?:[0-9a-f]{2})?$/i.test(color)
	);
}

/** The fill colour a picker shows: the shape's own, else the default fill. */
export function shapeFillColorOf(style: ShapeFillTarget): string {
	return normalizeHexColor(style?.fillColor, DEFAULT_FILL_COLOR);
}

/** The fill kind `style` paints with, as the fill select shows it. */
export function shapeFillModeOf(style: ShapeFillTarget): ShapeFillMode {
	if (style?.useBackgroundFill) {
		return 'background';
	}
	const mode = style?.fillMode;
	if (mode === 'none' || mode === 'gradient' || mode === 'image' || mode === 'pattern') {
		return mode;
	}
	if (mode === undefined && style?.fillColor === 'transparent') {
		return 'none';
	}
	return 'solid';
}

/**
 * The style change that switches `style` to fill kind `mode`. Each kind gets
 * the defaults it needs to paint: a gradient gets two stops (the current fill
 * colour to white) when it has none, a pattern its preset and second colour, a
 * picture its stretch mode. Leaving a slide background fill clears the flag.
 */
export function shapeFillModeChange(
	style: ShapeFillTarget,
	mode: ShapeFillMode,
): Partial<ShapeStyle> {
	const leaveBackground: Partial<ShapeStyle> = style?.useBackgroundFill
		? { useBackgroundFill: false }
		: {};
	const color = isUsableFillColor(style?.fillColor) ? {} : { fillColor: DEFAULT_FILL_COLOR };
	switch (mode) {
		case 'background':
			// The shape paints the slide background: its own fill is not drawn
			// and is not saved (`useBgFill`); the canvas paints the stage's
			// published slide background (`slide-background-fill.ts`).
			return { useBackgroundFill: true, fillMode: 'none', fillColor: 'transparent' };
		case 'none':
			return { ...leaveBackground, fillMode: 'none', fillColor: 'transparent' };
		case 'image':
			return {
				...leaveBackground,
				fillMode: 'image',
				fillImageMode: style?.fillImageMode || 'stretch',
			};
		case 'gradient':
			return {
				...leaveBackground,
				fillMode: 'gradient',
				fillGradientType: style?.fillGradientType || 'linear',
				fillGradientAngle: style?.fillGradientAngle ?? 90,
				fillGradientStops: shapeFillStopsOf(style),
			};
		case 'pattern':
			return {
				...leaveBackground,
				...color,
				fillMode: 'pattern',
				fillPatternPreset: style?.fillPatternPreset || DEFAULT_PATTERN_PRESET,
				fillPatternBackgroundColor: style?.fillPatternBackgroundColor || DEFAULT_PATTERN_BACKGROUND,
			};
		default:
			return { ...leaveBackground, ...color, fillMode: 'solid' };
	}
}

const modeIs =
	(...modes: ShapeFillMode[]) =>
	(style: ShapeFillTarget): boolean =>
		modes.includes(shapeFillModeOf(style));

/** The fill select (untitled: each pane heads it). */
export const SHAPE_FILL_MODE_SECTION: ShapeFillSection = {
	id: 'fillMode',
	fields: [
		{
			id: 'fillMode',
			kind: 'select',
			labelKey: 'pptx.table.fillMode',
			wide: true,
			testId: 'fill-mode',
			options: SHAPE_FILL_MODE_OPTIONS,
			read: shapeFillModeOf,
			write: (style, value) => shapeFillModeChange(style, String(value) as ShapeFillMode),
		},
	],
};

/**
 * The plain fields of the chosen fill kind: transparency for a filled shape,
 * the gradient type and (linear) angle, and picture tiling. A binding draws it
 * after its solid colour picker and before the gradient stops.
 */
export const SHAPE_FILL_OPTIONS_SECTION: ShapeFillSection = {
	id: 'fillOptions',
	fields: [
		{
			id: 'fillOpacity',
			kind: 'percent',
			labelKey: 'pptx.strokeEffects.fillOpacity',
			unit: '%',
			wide: true,
			testId: 'fill-opacity',
			visibleWhen: modeIs('solid', 'gradient', 'image', 'pattern'),
			read: (style) => (style?.fillOpacity ?? 1) * 100,
			write: (_style, value) => ({ fillOpacity: Number(value) / 100 }),
		},
		{
			id: 'fillGradientType',
			kind: 'select',
			labelKey: 'pptx.fillAdvanced.gradientType',
			wide: true,
			testId: 'fill-gradient-type',
			options: GRADIENT_TYPE_OPTIONS,
			visibleWhen: modeIs('gradient'),
			read: (style) => style?.fillGradientType || 'linear',
			write: (_style, value) => ({
				fillMode: 'gradient',
				fillGradientType: value === 'radial' ? 'radial' : 'linear',
			}),
		},
		{
			id: 'fillGradientAngle',
			kind: 'angle',
			labelKey: 'pptx.fillAdvanced.gradientAngle',
			unit: '°',
			step: 1,
			decimals: 0,
			wide: true,
			testId: 'fill-gradient-angle',
			visibleWhen: (style) =>
				shapeFillModeOf(style) === 'gradient' && (style?.fillGradientType || 'linear') === 'linear',
			read: (style) => style?.fillGradientAngle ?? 90,
			write: (_style, value) => ({ fillMode: 'gradient', fillGradientAngle: Number(value) }),
		},
		{
			id: 'fillImageMode',
			kind: 'select',
			labelKey: 'pptx.fillAdvanced.imageMode',
			wide: true,
			testId: 'fill-image-mode',
			options: IMAGE_MODE_OPTIONS,
			visibleWhen: modeIs('image'),
			read: (style) => style?.fillImageMode || 'stretch',
			write: (_style, value) => ({ fillImageMode: value === 'tile' ? 'tile' : 'stretch' }),
		},
	],
};

/** The fill select as a section list, what a generic renderer takes. */
export const SHAPE_FILL_MODE_SECTIONS: readonly ShapeFillSection[] = [SHAPE_FILL_MODE_SECTION];
/** The fill kind's plain fields as a section list. */
export const SHAPE_FILL_OPTIONS_SECTIONS: readonly ShapeFillSection[] = [
	SHAPE_FILL_OPTIONS_SECTION,
];
