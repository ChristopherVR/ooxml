import {
	TEXT_EFFECT_SECTIONS,
	buildTextWarpGallery,
	registerPptxWebControls,
	textWarpStyle,
} from 'ooxml-ui/pptx';
import type { RibbonGalleryPickEvent } from 'ooxml-ui/pptx';
import type { TextStyle } from 'pptx-viewer-core';

import type { Translator } from '../../i18n';
import { createInspectorCheckbox } from './controls-extra';
import { createInspectorFields } from './inspector-fields';
import type { InspectorHandlers, InspectorState } from './types';

export interface TextEffectsControls {
	el: HTMLElement;
	update(state: InspectorState): void;
}

export function createTextEffectsControls(
	doc: Document,
	t: Translator,
	handlers: InspectorHandlers,
): TextEffectsControls {
	registerPptxWebControls();
	const el = doc.createElement('div');
	el.className = 'pptxv-inspector-text-effects';
	let state: InspectorState | null = null;
	const apply = (patch: Partial<TextStyle>): void =>
		handlers.setTextStyle(patch, state?.selectedTextRange);
	const applyBody = (patch: Partial<TextStyle>): void => handlers.setTextStyle(patch);
	// The control is named explicitly: a wrapping `<label>` lends its whole text
	// content, which takes in an embedded number field's value too.
	const field = (labelText: string, input: HTMLElement): void => {
		const label = doc.createElement('label');
		label.textContent = labelText;
		input.setAttribute('aria-label', labelText);
		label.appendChild(input);
		el.appendChild(label);
	};
	const number = (label: string, key: keyof TextStyle, min = 0, body = false): HTMLInputElement => {
		const input = doc.createElement('input');
		input.type = 'number';
		input.min = String(min);
		input.addEventListener('change', () =>
			(body ? applyBody : apply)({ [key]: Number(input.value) }),
		);
		field(label, input);
		return input;
	};
	const color = (label: string, key: keyof TextStyle): HTMLInputElement => {
		const input = doc.createElement('input');
		input.type = 'color';
		input.addEventListener('input', () => apply({ [key]: input.value }));
		// B6: push into the "Recent colours" MRU list once the picker commits
		// (native `change`), never on the continuous `input` a drag fires.
		input.addEventListener('change', () => handlers.pushRecentColor(input.value));
		field(label, input);
		return input;
	};
	const toggle = (label: string, key: keyof TextStyle): HTMLInputElement => {
		const input = createInspectorCheckbox(doc);
		input.addEventListener('change', () => apply({ [key]: input.checked }));
		field(label, input);
		return input;
	};
	const columns = number(t('pptx.textFormatting.columns'), 'columnCount', 1, true);
	columns.max = '6';
	const columnSpacing = number(t('pptx.textFormatting.columnSpacing'), 'columnSpacing', 0, true);
	const outlineColor = color(t('pptx.textEffects.outline'), 'textOutlineColor');
	const outlineWidth = number(t('pptx.textEffects.outlineWidth'), 'textOutlineWidth');
	// Text shadow, glow and reflection: the shared descriptors (a toggle per
	// effect, its fields once on), applied like the other run effects here, to
	// the selected range when there is one.
	const effects = createInspectorFields(doc, t, {
		sections: TEXT_EFFECT_SECTIONS,
		onPatch: apply,
		onColorCommit: (hex) => handlers.pushRecentColor(hex),
	});
	el.appendChild(effects.el);
	const strike = toggle(t('pptx.textFormatting.strikethrough'), 'strikethrough');
	const highlight = color(t('pptx.text.highlightColor'), 'highlightColor');
	// Transform: the shared Text Warp gallery, one tile per preset. Captioned by a
	// legend, not a `<label>`: clicking a label's text would press its first tile
	// (No Transform) and clear the warp.
	const warpField = doc.createElement('fieldset');
	const warpCaption = doc.createElement('legend');
	warpCaption.textContent = t('pptx.textEffects.transform');
	const warp = doc.createElement('pptx-ui-ribbon-gallery');
	warp.setAttribute('mode', 'panel');
	warp.setAttribute('compact', '');
	warp.translateLabel = (key, params) => t(key, params ? { ...params } : undefined);
	warp.addEventListener('gallery-pick', (event) => {
		const style = textWarpStyle((event as RibbonGalleryPickEvent).detail.itemId);
		if (style) {
			applyBody(style);
		}
	});
	warpField.append(warpCaption, warp);
	el.appendChild(warpField);
	const warpAdjust = number(t('pptx.textEffects.adjustment'), 'textWarpAdj', 0, true);
	const inputs = Array.from(el.querySelectorAll<HTMLInputElement>('input,pptx-ui-checkbox'));

	return {
		el,
		update(next) {
			state = next;
			el.hidden = !next.canText;
			const style = next.textStyle ?? {};
			columns.value = String(style.columnCount ?? 1);
			columnSpacing.value = String(style.columnSpacing ?? 0);
			outlineColor.value = style.textOutlineColor ?? '#000000';
			outlineWidth.value = String(style.textOutlineWidth ?? 0);
			effects.update(next.textStyle, !next.canText);
			strike.checked = Boolean(style.strikethrough);
			highlight.value = style.highlightColor ?? '#ffff00';
			warp.descriptor = buildTextWarpGallery(style.textWarpPreset, !next.canText);
			warp.disabled = !next.canText;
			warpAdjust.value = String(style.textWarpAdj ?? 50000);
			for (const input of inputs) {
				input.disabled = !next.canText;
			}
		},
	};
}
