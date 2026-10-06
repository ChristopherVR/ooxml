/**
 * SmartArt Design > Layouts: the layout families the core SmartArt model can
 * switch between (`SWITCHABLE_LAYOUT_TYPES`, the same set the Inspector's layout
 * switcher offers). A pick runs the same `switchSmartArtLayout` the Inspector
 * runs, keeping the nodes, connections and styling.
 *
 * @module render/ribbon-galleries/smartart-layouts-gallery
 */
import { SWITCHABLE_LAYOUT_TYPES } from 'pptx-viewer-core';

import { SMARTART_LAYOUT_LABEL_KEYS } from '../schema-label-keys';
import type { RibbonGalleryModule } from './gallery-module';
import { galleryColorScheme } from './gallery-theme';
import { smartArtElementPatch, smartArtLayoutSwitchPatch } from './smartart-gallery-patch';
import { smartArtLayoutTileSvg } from './smartart-layout-tiles';

const TILE = { width: 56, height: 36 };
const humanize = (id: string): string => id.replace(/^./u, (c) => c.toUpperCase());

export const SMARTART_LAYOUTS_GALLERY: RibbonGalleryModule = {
	build(ctx) {
		const element = ctx.element;
		const data = element?.type === 'smartArt' ? element.smartArtData : undefined;
		const current = data?.resolvedLayoutType ?? 'list';
		const accent1 = galleryColorScheme(ctx).accent1;
		return {
			id: 'smartArtLayouts',
			labelKey: 'pptx.gallery.smartArtLayouts.title',
			label: 'Layouts',
			disabled: !data,
			sections: [
				{
					id: 'layouts',
					columns: 5,
					tileWidth: TILE.width,
					tileHeight: TILE.height,
					items: SWITCHABLE_LAYOUT_TYPES.map((type) => ({
						id: type,
						labelKey: SMARTART_LAYOUT_LABEL_KEYS[type] ?? `pptx.smartart.category.${type}`,
						label: humanize(type),
						previewSvg: smartArtLayoutTileSvg(type, accent1, TILE),
						applied: current === type,
					})),
				},
			],
		};
	},
	apply(itemId, ctx) {
		const element = ctx.element;
		const type = SWITCHABLE_LAYOUT_TYPES.find((candidate) => candidate === itemId);
		if (!type || element?.type !== 'smartArt' || !element.smartArtData) {
			return null;
		}
		if ((element.smartArtData.resolvedLayoutType ?? 'list') === type) {
			return null;
		}
		// drawingShapes is forwarded (cleared) so the reflow regenerates the new layout.
		const patch = smartArtElementPatch(
			element,
			smartArtLayoutSwitchPatch(element.smartArtData, type),
		);
		return patch ? { kind: 'element', ...patch } : null;
	},
};
