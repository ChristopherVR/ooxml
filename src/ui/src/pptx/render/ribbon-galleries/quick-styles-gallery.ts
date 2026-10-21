/**
 * The inspector's Quick Styles picker: every `SHAPE_QUICK_STYLES` preset (solid fills, outlines,
 * fills with a shadow) as one tile drawn by `shapeTileSvg`, in rows of six. A pick merges the
 * preset over the shape's style (`shapeQuickStylePatch`), so a style the preset does not set is
 * kept. No tile is marked: a preset is an overlay, so a shape's style cannot tell which one, if
 * any, was applied last.
 *
 * @module render/ribbon-galleries/quick-styles-gallery
 */
import type { PptxElement, ShapeStyle } from 'ooxml-core/pptx';
import { hasShapeProperties } from 'ooxml-core/pptx';

import { SHAPE_QUICK_STYLES, shapeQuickStylePatch } from '../shape-quick-styles';
import type { ShapeQuickStyle } from '../shape-quick-styles';
import type { RibbonGalleryModule } from './gallery-module';
import { shapeTileSvg, tileSpecFromShapeStyle } from './gallery-preview-svg';
import type { RibbonGalleryDescriptor } from './gallery-types';

const TILE = { width: 34, height: 22 };

/** A preset's tile id: the last segment of its label key (`blueFill`). */
function quickStyleId(preset: ShapeQuickStyle): string {
	return preset.labelKey.slice(preset.labelKey.lastIndexOf('.') + 1);
}

/** The Quick Styles gallery; `disabled` when the selection takes no shape style. */
export function buildQuickStylesGallery(disabled = false): RibbonGalleryDescriptor {
	return {
		id: 'quickStyles',
		labelKey: 'pptx.shape.quickStyles',
		label: 'Quick Styles',
		disabled,
		sections: [
			{
				id: 'presets',
				columns: 6,
				tileWidth: TILE.width,
				tileHeight: TILE.height,
				items: SHAPE_QUICK_STYLES.map((preset) => {
					const id = quickStyleId(preset);
					return {
						id,
						labelKey: preset.labelKey,
						label: preset.name,
						previewSvg: shapeTileSvg({
							...tileSpecFromShapeStyle(`quick-${id}`, preset.style, TILE),
							inset: 4,
						}),
						applied: false,
					};
				}),
			},
		],
	};
}

/** The shape style a Quick Styles tile sets, for hosts that merge it themselves. */
export function quickStyleShapeStyle(itemId: string): Partial<ShapeStyle> | undefined {
	return SHAPE_QUICK_STYLES.find((preset) => quickStyleId(preset) === itemId)?.style;
}

/** The selection, when it takes a shape style. */
function styledElement(element: PptxElement | null) {
	return element && hasShapeProperties(element) ? element : undefined;
}

export const QUICK_STYLES_GALLERY: RibbonGalleryModule = {
	build(ctx) {
		return buildQuickStylesGallery(!styledElement(ctx.element));
	},
	apply(itemId, ctx) {
		const element = styledElement(ctx.element);
		const style = quickStyleShapeStyle(itemId);
		return element && style
			? { kind: 'element', elementId: element.id, patch: shapeQuickStylePatch(element, style) }
			: null;
	},
};
