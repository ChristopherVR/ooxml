import { resolveObjectURL } from 'node:buffer';

import type { MediaPptxElement, PptxSlide } from 'ooxml-core/pptx';
import { describe, expect, it, vi } from 'vitest';

import type { MediaArrayBufferSource } from './media-element-source';
import { resolveMediaUrls } from './resolve-media-urls';

function mediaElement(overrides: Partial<MediaPptxElement> = {}): MediaPptxElement {
	return {
		id: 'm1',
		type: 'media',
		x: 0,
		y: 0,
		width: 100,
		height: 100,
		mediaType: 'video',
		...overrides,
	} as MediaPptxElement;
}

function slide(overrides: Record<string, unknown>): PptxSlide {
	return { id: 's1', elements: [], ...overrides } as unknown as PptxSlide;
}

function archive(parts: Record<string, string>): MediaArrayBufferSource {
	return {
		getMediaArrayBuffer: vi.fn(async (path: string) =>
			path in parts ? new TextEncoder().encode(parts[path]).buffer : undefined,
		),
		getImageData: vi.fn(async () => undefined),
	};
}

async function blobText(url: string | undefined): Promise<string | undefined> {
	return url === undefined ? undefined : resolveObjectURL(url)?.text();
}

describe('resolveMediaUrls', () => {
	it('gives embedded transition and click-effect sounds a playable Blob URL', async () => {
		const handler = archive({ 'ppt/media/audio1.wav': 'whoosh', 'ppt/media/audio2.wav': 'chime' });
		const blobUrls: string[] = [];
		const slides = [
			slide({ transition: { type: 'fade', soundPath: 'ppt/media/audio1.wav' } }),
			slide({ nativeAnimations: [{ soundPath: 'ppt/media/audio2.wav' }] }),
		];

		const urls = await resolveMediaUrls(handler, slides, blobUrls);

		expect(await blobText(urls.get('ppt/media/audio1.wav'))).toBe('whoosh');
		expect(await blobText(urls.get('ppt/media/audio2.wav'))).toBe('chime');
		expect(blobUrls).toHaveLength(2);
		for (const url of blobUrls) {
			URL.revokeObjectURL(url);
		}
	});

	it('leaves a linked sound out of the map, so consent still decides', async () => {
		const handler = archive({});
		const linked = 'https://media.example.com/a.wav';

		const urls = await resolveMediaUrls(
			handler,
			[slide({ transition: { type: 'fade', soundPath: linked } })],
			[],
		);

		expect(urls.has(linked)).toBe(false);
		expect(handler.getMediaArrayBuffer).not.toHaveBeenCalled();
	});

	it('hands a linked media element its own URL without reading the archive', async () => {
		const handler = archive({});
		const media = mediaElement({ mediaPath: 'https://cdn.example.com/demo.mp4' });
		const blobUrls: string[] = [];

		const urls = await resolveMediaUrls(handler, [slide({ elements: [media] })], blobUrls);

		expect(urls.get('https://cdn.example.com/demo.mp4')).toBe('https://cdn.example.com/demo.mp4');
		expect(media.mediaMissing).not.toBeTruthy();
		expect(handler.getMediaArrayBuffer).not.toHaveBeenCalled();
		expect(blobUrls).toHaveLength(0);
	});

	it('marks an embedded element missing when the archive lookup fails', async () => {
		const media = mediaElement({ mediaPath: 'ppt/media/missing.mp4' });

		await resolveMediaUrls(archive({}), [slide({ elements: [media] })], []);

		expect(media.mediaMissing).toBeTruthy();
	});
});
