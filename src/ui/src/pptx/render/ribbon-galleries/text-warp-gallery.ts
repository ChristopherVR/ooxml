/**
 * The inspector's Text Warp (WordArt Transform) picker: every OOXML `a:prstTxWarp` preset of
 * `TEXT_WARP_PRESETS`, one tile each drawn from `warpPreviewPath`. A pick sets the text style's
 * `textWarpPreset`; No Transform clears it.
 *
 * @module render/ribbon-galleries/text-warp-gallery
 */
import type { PptxElement, PptxTextWarpPreset, TextStyle } from 'ooxml-core/pptx';
import { hasTextProperties } from 'ooxml-core/pptx';

import { TEXT_WARP_PRESETS, warpPreviewPath } from '../text-warp-presets';
import type { RibbonGalleryModule } from './gallery-module';
import type { RibbonGalleryDescriptor } from './gallery-types';

/** The element whose text the gallery warps, when the selection has text. */
function textElement(element: PptxElement | null) {
	return element && hasTextProperties(element) ? element : undefined;
}

/** A fixed line-art tile; built from the catalogue only, never from user text. */
function tile(preset: PptxTextWarpPreset): string {
	return `<svg xmlns="http://www.w3.org/2000/svg" width="40" height="20" viewBox="0 0 40 20" aria-hidden="true"><path d="${warpPreviewPath(preset)}" stroke="currentColor" stroke-width="1.5" fill="none"/></svg>`;
}

/** The Text Warp gallery for text warped by `current` (none is No Transform). */
export function buildTextWarpGallery(
	current: PptxTextWarpPreset | undefined,
	disabled = false,
): RibbonGalleryDescriptor {
	const applied = current ?? 'textNoShape';
	return {
		id: 'textWarp',
		labelKey: 'pptx.textWarp.title',
		label: 'Text Warp',
		disabled,
		sections: [
			{
				id: 'presets',
				columns: 5,
				tileWidth: 40,
				tileHeight: 20,
				items: TEXT_WARP_PRESETS.map((preset) => ({
					id: preset.value,
					labelKey: preset.labelKey,
					label: preset.label,
					previewSvg: tile(preset.value),
					applied: preset.value === applied,
				})),
			},
		],
	};
}

/**
 * The text style change a Text Warp tile makes: its preset, or none for No Transform. Undefined
 * for an id that is not a preset.
 */
export function textWarpStyle(itemId: string): Pick<TextStyle, 'textWarpPreset'> | undefined {
	const preset = TEXT_WARP_PRESETS.find((entry) => entry.value === itemId)?.value;
	return preset ? { textWarpPreset: preset === 'textNoShape' ? undefined : preset } : undefined;
}

export const TEXT_WARP_GALLERY: RibbonGalleryModule = {
	build(ctx) {
		const element = textElement(ctx.element);
		return buildTextWarpGallery(element?.textStyle?.textWarpPreset, !element);
	},
	apply(itemId, ctx) {
		const element = textElement(ctx.element);
		const style = textWarpStyle(itemId);
		return element && style
			? {
					kind: 'element',
					elementId: element.id,
					patch: { textStyle: { ...element.textStyle, ...style } } as Partial<PptxElement>,
				}
			: null;
	},
};
