/**
 * `ribbon-home-commands` - the Home tab's three structurally-dead commands
 * (Reset, Shape Fill, Shape Outline) expressed as pure decision functions.
 *
 * WHY shared: React and Vue shipped all three as controls with the right label
 * and the right enabled state and no effect whatsoever. React's swatches wrote
 * `{ fill }` / `{ outlineColor }`, which are not fields of `ShapeStyle`
 * (`fillColor` / `strokeColor` are), so the buttons would still have done
 * nothing had the handler ever been passed down. Naming the change once here
 * is what stops a sixth spelling of the same two keys appearing.
 *
 * @module render/ribbon-home-commands
 */
import type { PptxSlide, PptxThemeColorRef, ShapeStyle } from 'ooxml-core/pptx';

import { DEFAULT_STROKE_WIDTH } from './arrange-extras';

/**
 * The layout PowerPoint's Home > Reset re-applies to the active slide,
 * restoring inherited placeholder geometry and formatting. Returns undefined
 * when the slide records no layout, in which case Reset has nothing to do and
 * the control must be inert rather than destructive.
 */
export function resetSlideLayoutPath(slide: PptxSlide | undefined): string | undefined {
	const path = slide?.layoutPath;
	return typeof path === 'string' && path !== '' ? path : undefined;
}

/**
 * The style change Home > Shape Fill commits for a picked swatch. Pass `ref`
 * for a theme-swatch pick (wins on save, so the fill follows a later theme
 * change); omit it (or pass `undefined`) to explicitly clear a previously-
 * stored ref for a plain/custom/recent pick.
 */
export function shapeFillChange(hex: string, ref?: PptxThemeColorRef): Partial<ShapeStyle> {
	return { fillColor: hex, fillColorRef: ref, fillMode: 'solid' };
}

/**
 * The style change an outline colour pick commits (Home > Shape Outline and
 * the inspector's line colour). Same `ref` contract as {@link shapeFillChange}.
 *
 * Pass the element's `current` style: a shape with no visible line (no width,
 * or a zero width from `a:noFill`) only got the colour, so nothing appeared.
 * Such a pick also turns the line on with {@link DEFAULT_STROKE_WIDTH} (1px,
 * which is PowerPoint's default 0.75pt line) and a solid line fill, the way
 * PowerPoint does; a line that already shows keeps its width and fill.
 */
export function shapeOutlineChange(
	hex: string,
	ref?: PptxThemeColorRef,
	current?: Pick<ShapeStyle, 'strokeWidth'>,
): Partial<ShapeStyle> {
	const change: Partial<ShapeStyle> = { strokeColor: hex, strokeColorRef: ref };
	const lineShows = (current?.strokeWidth ?? 0) > 0;
	if (current && !lineShows) {
		change.strokeWidth = DEFAULT_STROKE_WIDTH;
		change.strokeFillMode = 'solid';
	}
	return change;
}

/**
 * The swatch grid both pickers show: PowerPoint's Standard Colors row plus
 * black/white, which is what the vanilla binding already offered and what the
 * other four each re-listed differently.
 */
export const RIBBON_SHAPE_SWATCHES: readonly string[] = [
	'#ffffff',
	'#000000',
	'#c00000',
	'#ff0000',
	'#ffc000',
	'#ffff00',
	'#92d050',
	'#00b050',
	'#00b0f0',
	'#0070c0',
	'#002060',
	'#7030a0',
];
