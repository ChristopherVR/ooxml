/**
 * SmartArt Design > SmartArt Styles: Office's full built-in quick-style catalogue (the five Best
 * Match for Document styles and the nine 3-D styles; `DIAGRAM_NATIVE_STYLE_ENTRIES`). A pick sets
 * the nearest model intensity (`smartArtData.style`) for the editor's rendering and records the
 * built-in definition, which save writes as the quick-style part with its 3D properties (see
 * `smartart-native-entries.ts`). A model intensity (`flat`) still applies its definition.
 *
 * @module render/ribbon-galleries/smartart-styles-gallery
 */
import { DIAGRAM_NATIVE_STYLE_ENTRIES, DIAGRAM_NATIVE_STYLE_SECTIONS } from 'ooxml-core/diagram';

import { smartArtStyleDefinitionTileSvg } from '../../../smartart/definition-tiles';
import type { RibbonGalleryModule } from './gallery-module';
import { galleryColorScheme } from './gallery-theme';
import { smartArtElementPatch } from './smartart-gallery-patch';
import {
	appliedSmartArtStyle,
	builtinDefinitionXml,
	modelStyleIntensity,
	nativeStyleOf,
	smartArtStyleLabelKey,
	smartArtStyleSectionKey,
} from './smartart-native-entries';

const TILE = { width: 56, height: 36 };

export const SMARTART_STYLES_GALLERY: RibbonGalleryModule = {
	build(ctx) {
		const element = ctx.element;
		const data = element?.type === 'smartArt' ? element.smartArtData : undefined;
		const intensity = data?.quickStyle?.effectIntensity;
		const current = data
			? (appliedSmartArtStyle(data) ??
				(intensity === 'intense' ? 'simple5' : intensity === 'moderate' ? 'simple4' : 'simple1'))
			: undefined;
		const accent1 = galleryColorScheme(ctx).accent1;
		return {
			id: 'smartArtStyles',
			labelKey: 'pptx.gallery.smartArtStyles.title',
			label: 'SmartArt Styles',
			disabled: !data,
			sections: DIAGRAM_NATIVE_STYLE_SECTIONS.map((section) => ({
				id: section.id === '3D' ? 'threeD' : 'bestMatch',
				titleKey: smartArtStyleSectionKey(section.id),
				title: section.title,
				columns: 5,
				tileWidth: TILE.width,
				tileHeight: TILE.height,
				items: DIAGRAM_NATIVE_STYLE_ENTRIES.filter((entry) => entry.category === section.id).map(
					(entry) => ({
						id: entry.id,
						labelKey: smartArtStyleLabelKey(entry),
						label: entry.name,
						previewSvg: smartArtStyleDefinitionTileSvg(entry, accent1, TILE),
						applied: current === entry.id,
					}),
				),
			})),
		};
	},
	apply(itemId, ctx) {
		const entry = nativeStyleOf(itemId);
		if (!entry) return null;
		const patch = smartArtElementPatch(ctx.element, {
			style: modelStyleIntensity(entry),
			builtinQuickStyleXml: builtinDefinitionXml('style', entry.uniqueId),
		});
		return patch ? { kind: 'element', ...patch } : null;
	},
};
