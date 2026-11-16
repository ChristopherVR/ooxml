import {
	ARTISTIC_EFFECTS,
	ELEMENT_ALT_TEXT_SECTIONS,
	PICTURE_ADJUSTMENT_SECTIONS,
	PICTURE_CROP_SECTIONS,
} from 'ooxml-ui/pptx';

import type { Translator } from '../../i18n';
import { createEl } from '../../render';
import { makeRangeField, createInspectorSelect, createInspectorCheckbox } from './controls-extra';
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
 * renderer; the compound picture controls (artistic effect, duotone, colour
 * wash, replace and reset) stay native DOM here.
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

	const pct = (value: number): string => `${Math.round(value)}%`;

	const onPatch = (patch: Parameters<InspectorHandlers['patchSelectedElement']>[0]): void =>
		handlers.patchSelectedElement(patch);
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

	const artisticLabel = doc.createElement('label');
	artisticLabel.textContent = t('pptx.image.artisticEffects');
	const artistic = createInspectorSelect(doc);
	// Named explicitly: the wrapping `<label>` would otherwise lend the select
	// its whole text content, which includes every effect name.
	artistic.setAttribute('aria-label', t('pptx.image.artisticEffects'));
	for (const [value, labelKey] of ARTISTIC_EFFECTS) {
		const option = doc.createElement('option');
		option.value = value;
		option.textContent = t(labelKey);
		artistic.appendChild(option);
	}
	artistic.addEventListener('change', () =>
		handlers.setImageEffects({
			artisticEffect: artistic.value === 'none' ? undefined : artistic.value,
		}),
	);
	artisticLabel.appendChild(artistic);
	const duotone = createEl(doc, 'div', 'pptxv-inspector-grid');
	const colorInputs: Partial<Record<'color1' | 'color2', HTMLInputElement>> = {};
	const duoField = (labelText: string, key: 'color1' | 'color2'): HTMLInputElement => {
		const label = doc.createElement('label');
		label.textContent = labelText;
		const input = doc.createElement('input');
		input.type = 'color';
		input.addEventListener('input', () =>
			handlers.setImageEffects({
				duotone: {
					color1: key === 'color1' ? input.value : (colorInputs.color1?.value ?? '#000000'),
					color2: key === 'color2' ? input.value : (colorInputs.color2?.value ?? '#ffffff'),
				},
			}),
		);
		label.appendChild(input);
		duotone.appendChild(label);
		return input;
	};
	const color1 = duoField(t('pptx.image.duotoneDark'), 'color1');
	const color2 = duoField(t('pptx.image.duotoneLight'), 'color2');
	colorInputs.color1 = color1;
	colorInputs.color2 = color2;
	// B6: push into the "Recent colours" MRU list once each picker commits.
	for (const input of [color1, color2]) {
		input.addEventListener('change', () => handlers.pushRecentColor(input.value));
	}
	el.append(artisticLabel, duotone);
	const wash = createInspectorCheckbox(doc);
	const washColor = doc.createElement('input');
	washColor.type = 'color';
	let washOpacityValue = 40;
	const washOpacity = makeRangeField(doc, {
		label: t('pptx.image.washOpacity'),
		min: 0,
		max: 100,
		formatValue: pct,
		onCommit: (opacity) => {
			washOpacityValue = opacity;
			handlers.setImageEffects({ colorWash: { color: washColor.value, opacity } });
		},
	});
	wash.addEventListener('change', () =>
		handlers.setImageEffects({
			colorWash: wash.checked ? { color: washColor.value, opacity: washOpacityValue } : undefined,
		}),
	);
	washColor.addEventListener('input', () =>
		handlers.setImageEffects({
			colorWash: { color: washColor.value, opacity: washOpacityValue },
		}),
	);
	// B6: push into the "Recent colours" MRU list once the wash picker commits.
	washColor.addEventListener('change', () => handlers.pushRecentColor(washColor.value));
	el.append(wash, washColor, washOpacity.el);

	el.append(crop.el, resetCrop, altText.el);

	return {
		el,
		update(state) {
			el.hidden = !state.hasSelection || !state.isImage;
			const element = state.selectedElement;
			if (element) {
				adjustments.update(element, !state.isImage, element.id);
				// G7: `a:picLocks/@noCrop` forbids cropping this specific picture,
				// even on an otherwise-editable image.
				crop.update(element, !state.isImage || !state.croppable, element.id);
				altText.update(element, !state.isImage, element.id);
			}
			artistic.value = state.imageArtisticEffect;
			color1.value = state.imageDuotone1;
			color2.value = state.imageDuotone2;
			wash.checked = Boolean(state.imageColorWash);
			washColor.value = state.imageColorWash?.color ?? '#0066cc';
			washOpacity.setValue(state.imageColorWash?.opacity ?? 40);
			washOpacityValue = state.imageColorWash?.opacity ?? 40;
			artistic.disabled = !state.isImage;
			color1.disabled = !state.isImage;
			color2.disabled = !state.isImage;
			replace.disabled = !state.isImage;
			reset.disabled = !state.isImage || !state.imageResettable;
			resetCrop.disabled = !state.isImage || !state.croppable;
			wash.disabled = !state.isImage;
			washColor.disabled = !state.isImage;
			washOpacity.setDisabled(!state.isImage);
		},
	};
}
