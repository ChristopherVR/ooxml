import type { MediaPptxElement } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { translationsEn } from '../../i18n/translations-en';
import { inspectorFieldPatch, resolveInspectorSections } from './inspector-field-model';
import type { MediaTrimTarget } from './media-trim-fields';
import { MEDIA_TRIM_SECTIONS } from './media-trim-fields';

const target = (media: Partial<MediaPptxElement>, durationMs: number): MediaTrimTarget => ({
	media: { type: 'media', id: 'm1', x: 0, y: 0, width: 1, height: 1, ...media } as MediaPptxElement,
	durationMs,
});

const patch = (t: MediaTrimTarget, field: string, raw: string) =>
	inspectorFieldPatch(MEDIA_TRIM_SECTIONS, 'mediaTrim', field, t, raw);

describe('media trim fields', () => {
	it('shows the start only until the duration is known', () => {
		const [view] = resolveInspectorSections(MEDIA_TRIM_SECTIONS, target({ trimStartMs: 1500 }, 0));
		expect(view!.fields.map((f) => [f.id, f.value, f.max])).toStrictEqual([
			['trimStart', 1.5, undefined],
		]);
	});

	it('shows the end as an absolute time bounded by the live duration', () => {
		const [view] = resolveInspectorSections(
			MEDIA_TRIM_SECTIONS,
			target({ trimStartMs: 1000, trimEndMs: 5000 }, 20000),
		);
		expect(view!.fields.map((f) => [f.id, f.value, f.max])).toStrictEqual([
			['trimStart', 1, 15],
			['trimEnd', 15, 20],
		]);
	});

	it('stores the end as the distance from the tail and keeps the start before it', () => {
		const t = target({ trimStartMs: 2000, trimEndMs: 5000 }, 20000);
		expect(patch(t, 'trimEnd', '12')).toStrictEqual({ trimEndMs: 8000 });
		expect(patch(t, 'trimEnd', '99')).toStrictEqual({ trimEndMs: 0 });
		expect(patch(t, 'trimEnd', '1')).toStrictEqual({ trimEndMs: 18000 });
		expect(patch(t, 'trimStart', '30')).toStrictEqual({ trimStartMs: 15000 });
	});

	it('labels every field with a dictionary key', () => {
		const keys = MEDIA_TRIM_SECTIONS[0]!.fields.map((f) => f.labelKey);
		expect(keys.filter((key) => !(key in translationsEn))).toStrictEqual([]);
	});
});
