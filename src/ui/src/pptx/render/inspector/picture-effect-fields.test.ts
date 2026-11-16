import type { PptxElement } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { translationsEn } from '../../i18n/translations-en';
import {
	inspectorFieldPatch,
	inspectorTogglePatch,
	resolveInspectorSections,
} from './inspector-field-model';
import { PICTURE_EFFECT_SECTIONS, pictureDuotonePresetPatch } from './picture-effect-fields';

const picture = (imageEffects?: Record<string, unknown>): PptxElement =>
	({ type: 'image', id: 'p1', x: 0, y: 0, width: 10, height: 10, imageEffects }) as PptxElement;

describe('picture colour effect fields', () => {
	it('shows each effect as an off toggle on an unadjusted picture', () => {
		const views = resolveInspectorSections(PICTURE_EFFECT_SECTIONS, picture());
		expect(views.map((v) => [v.id, v.toggleable, v.on, v.fields.length])).toStrictEqual([
			['pictureColorWash', true, false, 0],
			['pictureRecolor', true, false, 0],
			['pictureDuotone', true, false, 0],
		]);
	});

	it('switches an effect on with its defaults and keeps the others', () => {
		const el = picture({ brightness: 10 });
		expect(
			inspectorTogglePatch(PICTURE_EFFECT_SECTIONS, 'pictureColorWash', el, true),
		).toStrictEqual({
			imageEffects: { brightness: 10, colorWash: { color: '#0066cc', opacity: 40 } },
		});
		expect(
			inspectorTogglePatch(PICTURE_EFFECT_SECTIONS, 'pictureDuotone', el, false),
		).toStrictEqual({
			imageEffects: { brightness: 10, duotone: undefined },
		});
	});

	it('edits one colour of an effect and hides the target colour while transparent', () => {
		const el = picture({
			clrChange: { clrFrom: '#ffffff', clrTo: '#000000', clrToTransparent: true },
		});
		const [, recolor] = resolveInspectorSections(PICTURE_EFFECT_SECTIONS, el);
		expect(recolor!.fields.map((f) => f.id)).toStrictEqual(['clrFrom', 'clrToTransparent']);
		expect(
			inspectorFieldPatch(PICTURE_EFFECT_SECTIONS, 'pictureRecolor', 'clrFrom', el, '#123456'),
		).toStrictEqual({
			imageEffects: { clrChange: { clrFrom: '#123456', clrTo: '#000000', clrToTransparent: true } },
		});
		expect(
			pictureDuotonePresetPatch(picture(), { shadow: '#111111', highlight: '#eeeeee' }),
		).toStrictEqual({ imageEffects: { duotone: { color1: '#111111', color2: '#eeeeee' } } });
	});

	it('writes nothing to a non-picture and labels every field with a dictionary key', () => {
		const shape = { type: 'shape', id: 's1', x: 0, y: 0, width: 1, height: 1 } as PptxElement;
		expect(
			inspectorTogglePatch(PICTURE_EFFECT_SECTIONS, 'pictureDuotone', shape, true),
		).toStrictEqual({});
		const keys = PICTURE_EFFECT_SECTIONS.flatMap((s) => [
			s.labelKey ?? '',
			...s.fields.map((f) => f.labelKey),
		]);
		expect(keys.filter((key) => !(key in translationsEn))).toStrictEqual([]);
	});
});
