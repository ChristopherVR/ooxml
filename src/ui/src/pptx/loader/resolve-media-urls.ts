import type { MediaPptxElement, PptxSlide } from 'ooxml-core/pptx';

import { collectAnimationSoundPaths, collectMediaElements } from './load-content-helpers';
import { type MediaArrayBufferSource, resolveMediaElementSource } from './media-element-source';

/**
 * Build a deck's media URL map, the one every binding's load pipeline uses:
 * each media element's playable source (Blob URL, poster data URL, or a linked
 * element's own URL), plus a Blob URL for every embedded transition and
 * animation sound (`p:stSnd`), which backs no media element. The slide show
 * and presenter view play sounds only from this map, so a binding that built
 * it by hand and skipped the sounds played none.
 *
 * Linked sounds get no entry: `collectAnimationSoundPaths` leaves them out,
 * and `playDeckSound` plays their URL only once linked content is allowed.
 * Media elements with no playable source are marked `mediaMissing`. Each Blob
 * URL created is pushed onto `blobUrls`, which the caller revokes.
 */
export async function resolveMediaUrls(
	handler: MediaArrayBufferSource,
	slides: readonly PptxSlide[],
	blobUrls: string[],
): Promise<Map<string, string>> {
	const mediaElements: MediaPptxElement[] = [];
	for (const slide of slides) {
		collectMediaElements(slide.elements, mediaElements);
	}
	const urls = new Map<string, string>();
	await Promise.all(
		mediaElements.map(async (mediaElement) => {
			const resolved = await resolveMediaElementSource(mediaElement, handler);
			if (resolved.missing || !resolved.mediaPath || !resolved.url) {
				mediaElement.mediaMissing = true;
				return;
			}
			urls.set(resolved.mediaPath, resolved.url);
			if (resolved.isBlobUrl) {
				blobUrls.push(resolved.url);
			}
		}),
	);

	const soundPaths = collectAnimationSoundPaths(slides).filter((path) => !urls.has(path));
	await Promise.all(
		soundPaths.map(async (soundPath) => {
			try {
				const arrayBuffer = await handler.getMediaArrayBuffer(soundPath);
				if (arrayBuffer) {
					const blobUrl = URL.createObjectURL(new Blob([arrayBuffer]));
					blobUrls.push(blobUrl);
					urls.set(soundPath, blobUrl);
				}
			} catch {
				/* Non-critical: the sound simply will not play. */
			}
		}),
	);
	return urls;
}
