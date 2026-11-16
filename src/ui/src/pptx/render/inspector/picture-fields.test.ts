import type { PptxElement } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { translationsEn } from '../../i18n/translations-en';
import { ELEMENT_ALT_TEXT_SECTIONS } from './element-alt-text-fields';
import { inspectorFieldPatch, resolveInspectorSections } from './inspector-field-model';
import { PICTURE_ADJUSTMENT_SECTIONS, PICTURE_CROP_SECTIONS } from './picture-fields';

const picture = (extra: Record<string, unknown> = {}): PptxElement =>
	({ type: 'image', id: 'p1', x: 0, y: 0, width: 10, height: 10, ...extra }) as PptxElement;
const shape = (extra: Record<string, unknown> = {}): PptxElement =>
	({ type: 'shape', id: 's1', x: 0, y: 0, width: 10, height: 10, ...extra }) as PptxElement;

const adjust = (el: PptxElement, field: string, raw: string | boolean) =>
	inspectorFieldPatch(PICTURE_ADJUSTMENT_SECTIONS, 'pictureAdjustments', field, el, raw);
const crop = (el: PptxElement, field: string, raw: string) =>
	inspectorFieldPatch(PICTURE_CROP_SECTIONS, 'pictureCrop', field, el, raw);
const alt = (el: PptxElement, field: string, raw: string) =>
	inspectorFieldPatch(ELEMENT_ALT_TEXT_SECTIONS, 'altText', field, el, raw);

describe('picture adjustment fields', () => {
	it('reads neutral values for an unadjusted picture', () => {
		const [view] = resolveInspectorSections(PICTURE_ADJUSTMENT_SECTIONS, picture());
		expect(view!.fields.map((f) => [f.id, f.control, f.value, f.min, f.max])).toStrictEqual([
			['brightness', 'range', 0, -100, 100],
			['contrast', 'range', 0, -100, 100],
			['saturation', 'range', 0, -100, 100],
			['transparency', 'range', 0, 0, 100],
			['biLevel', 'range', 0, 0, 100],
			['grayscale', 'checkbox', false, undefined, undefined],
		]);
	});

	it('shows transparency inverted from the stored opacity', () => {
		const [view] = resolveInspectorSections(
			PICTURE_ADJUSTMENT_SECTIONS,
			picture({ imageEffects: { alphaModFix: 30 } }),
		);
		expect(view!.fields.find((f) => f.id === 'transparency')?.value).toBe(70);
	});

	it('merges each change into the picture effects', () => {
		const el = picture({
			imageEffects: { contrast: 5, duotone: { color1: '#000', color2: '#fff' } },
		});
		expect(adjust(el, 'brightness', '140')).toStrictEqual({
			imageEffects: { contrast: 5, duotone: { color1: '#000', color2: '#fff' }, brightness: 100 },
		});
		expect(adjust(el, 'transparency', '25')).toMatchObject({
			imageEffects: { contrast: 5, alphaModFix: 75 },
		});
		expect(adjust(el, 'grayscale', true)).toMatchObject({ imageEffects: { grayscale: true } });
	});

	it('clears a zero transparency, threshold or grayscale instead of writing it', () => {
		const el = picture({ imageEffects: { alphaModFix: 40, biLevel: 50, grayscale: true } });
		expect(adjust(el, 'transparency', '0')).toStrictEqual({
			imageEffects: { alphaModFix: undefined, biLevel: 50, grayscale: true },
		});
		expect(adjust(el, 'biLevel', '0')).toMatchObject({ imageEffects: { biLevel: undefined } });
		expect(adjust(el, 'grayscale', false)).toMatchObject({
			imageEffects: { grayscale: undefined },
		});
	});

	it('writes nothing to a non-picture', () => {
		expect(adjust(shape(), 'brightness', '10')).toStrictEqual({});
		expect(adjust(shape(), 'transparency', '10')).toStrictEqual({});
	});
});

describe('picture crop fields', () => {
	it('shows each inset as a percentage up to the shared clamp', () => {
		const [view] = resolveInspectorSections(PICTURE_CROP_SECTIONS, picture({ cropLeft: 0.125 }));
		expect(view!.fields.map((f) => [f.id, f.value, f.max])).toStrictEqual([
			['cropLeft', 13, 90],
			['cropTop', 0, 90],
			['cropRight', 0, 90],
			['cropBottom', 0, 90],
		]);
	});

	it('writes clamped fractions', () => {
		expect(crop(picture(), 'cropTop', '20')).toStrictEqual({ cropTop: 0.2 });
		expect(crop(picture(), 'cropRight', '99')).toStrictEqual({ cropRight: 0.9 });
		expect(crop(picture(), 'cropBottom', '')).toBeUndefined();
	});
});

describe('alt text fields', () => {
	it('offers a picture its description only, and other kinds a title too', () => {
		const [pictureView] = resolveInspectorSections(ELEMENT_ALT_TEXT_SECTIONS, picture());
		expect(pictureView!.fields.map((f) => [f.id, f.control])).toStrictEqual([
			['altText', 'textarea'],
		]);
		const [shapeView] = resolveInspectorSections(
			ELEMENT_ALT_TEXT_SECTIONS,
			shape({ altText: 'A logo', title: 'Logo' }),
		);
		expect(shapeView!.fields.map((f) => [f.id, f.control, f.value])).toStrictEqual([
			['altText', 'textarea', 'A logo'],
			['title', 'text', 'Logo'],
		]);
		expect(shapeView!.fields[0]!.placeholderKey).toBe(
			'pptx.elementAccessibility.altTextPlaceholder',
		);
	});

	it('writes the text as typed, an empty box included', () => {
		expect(alt(shape(), 'altText', '  A chart ')).toStrictEqual({ altText: '  A chart ' });
		expect(alt(shape({ title: 'Old' }), 'title', '')).toStrictEqual({ title: '' });
	});
});

describe('picture and alt text labels', () => {
	it('labels every field with a dictionary key', () => {
		const keys = [
			...PICTURE_ADJUSTMENT_SECTIONS,
			...PICTURE_CROP_SECTIONS,
			...ELEMENT_ALT_TEXT_SECTIONS,
		].flatMap((section) =>
			section.fields.flatMap((f) => [f.labelKey, ...(f.placeholderKey ? [f.placeholderKey] : [])]),
		);
		expect(keys.filter((key) => !(key in translationsEn))).toStrictEqual([]);
	});
});
