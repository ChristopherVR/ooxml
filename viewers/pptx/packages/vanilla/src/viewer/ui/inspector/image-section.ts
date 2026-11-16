import {
	DUOTONE_PRESETS,
	ELEMENT_ALT_TEXT_SECTIONS,
	PICTURE_ADJUSTMENT_SECTIONS,
	PICTURE_ARTISTIC_SECTIONS,
	PICTURE_CROP_SECTIONS,
	PICTURE_EFFECT_SECTIONS,
	pictureDuotonePresetPatch,
} from 'ooxml-ui/pptx';
import type { PptxElement } from 'pptx-viewer-core';

import type { Translator } from '../../i18n';
import { createEl } from '../../render';
import { createInspectorFields } from './inspector-fields';
import type { InspectorHandlers, InspectorState } from './types';

export interface ImageSection {
	el: HTMLElement;
	update(state: InspectorState): void;
}

/**
 * The Image section. The adjustments, the crop insets and the alt text are
 * the shared picture descriptors (`PICTURE_ADJUSTMENT_SECTIONS`,
 * `PICTURE_CROP_SECTIONS`, `ELEMENT_ALT_TEXT_SECTIONS`) drawn by the generic
 * renderer; the artistic effect gallery and the colour effects are shared
 * descriptors too; the duotone presets, replace and reset stay native DOM here.
 */
export function createImageSection(
	doc: Document,
	t: Translator,
	section: (label: string) => HTMLElement,
	handlers: InspectorHandlers,
): ImageSection {
	const el = section(t('pptx.inspector.image'));
	const replace = doc.createElement('button');
	replace.type = 'button';
	replace.textContent = t('pptx.image.replaceImage');
	replace.addEventListener('click', handlers.replaceImage);
	const reset = doc.createElement('button');
	reset.type = 'button';
	reset.textContent = t('pptx.image.resetImage');
	reset.addEventListener('click', handlers.resetImage);
	el.append(replace, reset);
	const resetCrop = doc.createElement('button');
	resetCrop.type = 'button';
	resetCrop.textContent = t('pptx.image.resetCrop');
	resetCrop.addEventListener('click', handlers.resetCrop);

	const onPatch = (patch: Partial<PptxElement>): void => handlers.patchSelectedElement(patch);
	const adjustments = createInspectorFields(doc, t, {
		sections: PICTURE_ADJUSTMENT_SECTIONS,
		onPatch,
	});
	const crop = createInspectorFields(doc, t, {
		sections: PICTURE_CROP_SECTIONS,
		onPatch,
		sectionClassName: 'pptxv-inspector-grid',
	});
	const altText = createInspectorFields(doc, t, { sections: ELEMENT_ALT_TEXT_SECTIONS, onPatch });
	el.appendChild(adjustments.el);

	// The artistic effect gallery: the shared descriptor, previewing the picture.
	const artistic = createInspectorFields(doc, t, { sections: PICTURE_ARTISTIC_SECTIONS, onPatch });
	// Colour wash, recolour and duotone: the shared descriptors; the duotone
	// preset swatches below apply the shared `pictureDuotonePresetPatch`.
	const colorEffects = createInspectorFields(doc, t, {
		sections: PICTURE_EFFECT_SECTIONS,
		onPatch,
		onColorCommit: (hex) => handlers.pushRecentColor(hex),
	});
	let latestElement: PptxElement | undefined;
	const presets = createEl(doc, 'div', 'pptxv-duotone-presets');
	const presetButtons = DUOTONE_PRESETS.map((preset) => {
		const swatch = doc.createElement('button');
		swatch.type = 'button';
		swatch.className = 'pptxv-duotone-preset';
		swatch.dataset['testid'] = 'picture-duotone-preset';
		swatch.title = t(preset.labelKey);
		swatch.setAttribute('aria-label', t(preset.labelKey));
		swatch.style.background = `linear-gradient(135deg, ${preset.shadow} 50%, ${preset.highlight} 50%)`;
		swatch.addEventListener('click', () => {
			if (latestElement) {
				onPatch(pictureDuotonePresetPatch(latestElement, preset));
			}
		});
		return swatch;
	});
	presets.append(...presetButtons);
	el.append(artistic.el, colorEffects.el, presets);

	el.append(crop.el, resetCrop, altText.el);

	return {
		el,
		update(state) {
			el.hidden = !state.hasSelection || !state.isImage;
			const element = state.selectedElement;
			latestElement = element;
			for (const swatch of presetButtons) {
				swatch.disabled = !state.isImage;
			}
			if (element) {
				adjustments.update(element, !state.isImage, element.id);
				// G7: `a:picLocks/@noCrop` forbids cropping this specific picture,
				// even on an otherwise-editable image.
				crop.update(element, !state.isImage || !state.croppable, element.id);
				altText.update(element, !state.isImage, element.id);
				colorEffects.update(element, !state.isImage, element.id);
				artistic.update(element, !state.isImage, element.id);
			}
			replace.disabled = !state.isImage;
			reset.disabled = !state.isImage || !state.imageResettable;
			resetCrop.disabled = !state.isImage || !state.croppable;
		},
	};
}
