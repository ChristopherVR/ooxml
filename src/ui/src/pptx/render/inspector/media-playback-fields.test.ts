import type { PptxElement } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { translationsEn } from '../../i18n/translations-en';
import { inspectorFieldPatch, resolveInspectorSections } from './inspector-field-model';
import { MEDIA_PLAYBACK_SECTIONS } from './media-playback-fields';

const media = (extra: Record<string, unknown> = {}): PptxElement =>
	({
		type: 'media',
		id: 'm1',
		x: 0,
		y: 0,
		width: 10,
		height: 10,
		mediaType: 'video',
		...extra,
	}) as PptxElement;

const patch = (el: PptxElement, field: string, raw: string | boolean) =>
	inspectorFieldPatch(MEDIA_PLAYBACK_SECTIONS, 'mediaPlayback', field, el, raw);

describe('media playback fields', () => {
	it('offers full screen for a video and play across slides for audio', () => {
		const ids = (el: PptxElement) =>
			resolveInspectorSections(MEDIA_PLAYBACK_SECTIONS, el)[0]!.fields.map((f) => f.id);
		expect(ids(media())).toContain('fullScreen');
		expect(ids(media())).not.toContain('playAcrossSlides');
		expect(ids(media({ mediaType: 'audio' }))).toContain('playAcrossSlides');
		expect(ids(media({ mediaType: 'audio' }))).not.toContain('fullScreen');
	});

	it('reads PowerPoint defaults and writes the media fields', () => {
		const [view] = resolveInspectorSections(MEDIA_PLAYBACK_SECTIONS, media({ volume: 0.5 }));
		expect(view!.fields.find((f) => f.id === 'volume')?.value).toBe(50);
		expect(view!.fields.find((f) => f.id === 'start')?.value).toBe('onClick');
		expect(patch(media(), 'volume', '80')).toStrictEqual({ volume: 0.8 });
		expect(patch(media(), 'start', 'auto')).toStrictEqual({ autoPlay: true, startMode: 'auto' });
		expect(patch(media({ mediaType: 'audio' }), 'playAcrossSlides', true)).toStrictEqual({
			playAcrossSlides: true,
		});
	});

	it('clears a switched-off flag, a zero fade and normal speed', () => {
		expect(patch(media({ loop: true }), 'loop', false)).toStrictEqual({ loop: undefined });
		expect(patch(media(), 'fadeInDuration', '0')).toStrictEqual({ fadeInDuration: undefined });
		expect(patch(media(), 'playbackSpeed', '1')).toStrictEqual({ playbackSpeed: undefined });
		expect(patch(media(), 'playbackSpeed', '9')).toStrictEqual({ playbackSpeed: 4 });
	});

	it('writes nothing to another element and labels every field', () => {
		const shape = { type: 'shape', id: 's', x: 0, y: 0, width: 1, height: 1 } as PptxElement;
		expect(patch(shape, 'loop', true)).toStrictEqual({});
		const keys = MEDIA_PLAYBACK_SECTIONS[0]!.fields.flatMap((f) => [
			f.labelKey,
			...(f.options ?? []).map((o) => o.labelKey),
		]);
		expect(keys.filter((key) => !(key in translationsEn))).toStrictEqual([]);
	});
});
