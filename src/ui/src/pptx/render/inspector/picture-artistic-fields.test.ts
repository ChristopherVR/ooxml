import type { PptxElement } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { translationsEn } from '../../i18n/translations-en';
import { inspectorFieldPatch, resolveInspectorSections } from './inspector-field-model';
import { PICTURE_ARTISTIC_SECTIONS } from './picture-artistic-fields';

const picture = (extra: Record<string, unknown> = {}): PptxElement =>
	({ type: 'image', id: 'p', x: 0, y: 0, width: 1, height: 1, ...extra }) as PptxElement;

describe('picture artistic effect gallery', () => {
	it('previews every effect on the picture through its CSS filter', () => {
		const [view] = resolveInspectorSections(
			PICTURE_ARTISTIC_SECTIONS,
			picture({ imageData: 'data:image/png;base64,AA', imageEffects: { artisticEffect: 'blur' } }),
		);
		const field = view!.fields[0]!;
		expect([field.control, field.value, field.previewImage]).toStrictEqual([
			'gallery',
			'blur',
			'data:image/png;base64,AA',
		]);
		expect(field.options!.find((o) => o.value === 'blur')?.previewFilter).toBe('blur(4px)');
		expect(field.options![0]).toStrictEqual({ value: 'none', labelKey: 'pptx.image.effectNone' });
	});

	it('writes the effect and clears it for None, keeping the other effects', () => {
		const el = picture({ imageEffects: { brightness: 10, artisticEffect: 'blur' } });
		const write = (raw: string) =>
			inspectorFieldPatch(PICTURE_ARTISTIC_SECTIONS, 'pictureArtistic', 'artisticEffect', el, raw);
		expect(write('sepia')).toStrictEqual({
			imageEffects: { brightness: 10, artisticEffect: 'sepia' },
		});
		expect(write('none')).toStrictEqual({
			imageEffects: { brightness: 10, artisticEffect: undefined },
		});
	});

	it('labels every choice with a dictionary key', () => {
		const keys = PICTURE_ARTISTIC_SECTIONS[0]!.fields.flatMap((f) => [
			f.labelKey,
			...(f.options ?? []).map((o) => o.labelKey),
		]);
		expect(keys.filter((key) => !(key in translationsEn))).toStrictEqual([]);
	});
});
