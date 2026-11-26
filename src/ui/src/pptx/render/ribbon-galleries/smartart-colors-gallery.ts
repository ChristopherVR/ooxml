/**
 * SmartArt Design > Change Colors: Office's full built-in colour catalogue (Primary Theme
 * Colors, Colorful, Accent 1 to 6; `DIAGRAM_NATIVE_COLOR_ENTRIES`). A pick recolours the shapes
 * through the shared inspector patch and records the built-in definition for save (see
 * `smartart-native-entries.ts`). A model scheme name (`monochromatic1`) still applies the
 * definition it stands for.
 *
 * @module render/ribbon-galleries/smartart-colors-gallery
 */
import { DIAGRAM_NATIVE_COLOR_ENTRIES, DIAGRAM_NATIVE_COLOR_SECTIONS } from 'ooxml-core/diagram';

import {
	smartArtColorDefinitionTileSvg,
	smartArtDefinitionFills,
} from '../../../smartart/definition-tiles';
import { flattenNodes } from '../smartart-layout-helpers';
import type { RibbonGalleryModule } from './gallery-module';
import { galleryColorScheme } from './gallery-theme';
import type { RibbonGallerySection } from './gallery-types';
import { smartArtColorPatch } from './smartart-color-patch';
import { smartArtElementPatch } from './smartart-gallery-patch';
import {
	appliedSmartArtColor,
	builtinDefinitionXml,
	modelColorScheme,
	nativeColorOf,
	smartArtColorLabel,
	smartArtColorSectionKey,
} from './smartart-native-entries';

const TILE = { width: 56, height: 36 };

export const SMARTART_COLORS_GALLERY: RibbonGalleryModule = {
	build(ctx) {
		const element = ctx.element;
		const data = element?.type === 'smartArt' ? element.smartArtData : undefined;
		const scheme = galleryColorScheme(ctx);
		// A deck whose colours part names no built-in definition: match its parsed node fills.
		const fills = data?.colorTransform?.fillColors?.map((color) => color.toLowerCase());
		const matched = fills?.length
			? DIAGRAM_NATIVE_COLOR_ENTRIES.find(
					(entry) =>
						entry.preview.fills.length === fills.length &&
						smartArtDefinitionFills(entry, { ...scheme }, fills.length).every(
							(color, index) => color.toLowerCase() === fills[index],
						),
				)?.id
			: undefined;
		const current = data
			? (appliedSmartArtColor(data) ?? matched ?? (fills?.length ? undefined : 'colorful1'))
			: undefined;
		const sections: RibbonGallerySection[] = DIAGRAM_NATIVE_COLOR_SECTIONS.map((section) => ({
			id: section.id,
			titleKey: smartArtColorSectionKey(section.id),
			title: section.title,
			columns: 5,
			tileWidth: TILE.width,
			tileHeight: TILE.height,
			items: DIAGRAM_NATIVE_COLOR_ENTRIES.filter((entry) => entry.category === section.id).map(
				(entry) => {
					const { labelKey, labelParams } = smartArtColorLabel(entry);
					return {
						id: entry.id,
						labelKey,
						...(labelParams && { labelParams }),
						label: entry.name,
						previewSvg: smartArtColorDefinitionTileSvg(entry, { ...scheme }, TILE),
						applied: current === entry.id,
					};
				},
			),
		}));
		return {
			id: 'smartArtColors',
			labelKey: 'pptx.gallery.smartArtColors.title',
			label: 'Change Colors',
			disabled: !data,
			sections,
		};
	},
	apply(itemId, ctx) {
		const entry = nativeColorOf(itemId);
		const data = ctx.element?.type === 'smartArt' ? ctx.element.smartArtData : undefined;
		if (!entry || !data) return null;
		const fills = smartArtDefinitionFills(
			entry,
			{ ...galleryColorScheme(ctx) },
			Math.max(1, flattenNodes(data.nodes).length),
		);
		const xml = builtinDefinitionXml('colors', entry.uniqueId);
		const patch = smartArtElementPatch(ctx.element, {
			colorScheme: modelColorScheme(entry),
			...smartArtColorPatch(data, fills, entry.uniqueId),
			// Until the SmartArt library has loaded, save merges the fills into the existing part.
			builtinColorsXml: xml,
		});
		return patch ? { kind: 'element', ...patch } : null;
	},
};
