/**
 * A media element for an audio or video file a host's own picker returned, sized and centred as
 * Insert > Media does it, for every building-blocks port whose editor has no file-based media
 * insert of its own. DOM only (`FileReader`, a probing `<video>`).
 *
 * @module render/building-blocks/media-element-from-file
 */
import { createEditorId } from 'ooxml-core/pptx';
import type { PptxElement } from 'ooxml-core/pptx';

import type { CanvasSize } from '../../types';
import { classifyMediaType } from '../media-file-type';

/** Fixed size (px) of an inserted audio player strip. */
export const MEDIA_AUDIO_SIZE = { width: 420, height: 64 } as const;
/** Default and cap (px) of an inserted video; the aspect is kept when it can be probed. */
export const MEDIA_VIDEO_MAX_SIZE = { width: 640, height: 360 } as const;

/** A video data URL's natural size, capped to {@link MEDIA_VIDEO_MAX_SIZE}. */
export function probeVideoSize(
	doc: Document,
	dataUrl: string,
): Promise<{ width: number; height: number }> {
	const max = MEDIA_VIDEO_MAX_SIZE;
	return new Promise((resolve) => {
		const probe = doc.createElement('video');
		probe.preload = 'metadata';
		probe.onloadedmetadata = () => {
			const width = probe.videoWidth || max.width;
			const height = probe.videoHeight || max.height;
			if (width > max.width || height > max.height) {
				const scale = Math.min(max.width / width, max.height / height);
				resolve({ width: Math.round(width * scale), height: Math.round(height * scale) });
				return;
			}
			resolve({ width, height });
		};
		probe.onerror = () => resolve({ width: max.width, height: max.height });
		probe.src = dataUrl;
	});
}

/**
 * A centred media element for `file`, read as a data URL. Resolves null for a file that is not
 * audio or video, or that cannot be read.
 */
export function mediaElementFromFile(
	file: File,
	canvasSize: CanvasSize,
	doc: Document = document,
): Promise<PptxElement | null> {
	const mediaType = classifyMediaType(file.type);
	if (!mediaType) {
		return Promise.resolve(null);
	}
	return new Promise((resolve) => {
		const reader = new FileReader();
		reader.onload = async () => {
			const dataUrl = typeof reader.result === 'string' ? reader.result : '';
			if (!dataUrl) {
				resolve(null);
				return;
			}
			const size =
				mediaType === 'audio' ? { ...MEDIA_AUDIO_SIZE } : await probeVideoSize(doc, dataUrl);
			resolve({
				id: createEditorId('media'),
				type: 'media',
				mediaType,
				mediaMimeType: file.type || undefined,
				mediaData: dataUrl,
				x: Math.max(0, Math.round((canvasSize.width - size.width) / 2)),
				y: Math.max(0, Math.round((canvasSize.height - size.height) / 2)),
				width: size.width,
				height: size.height,
			} as unknown as PptxElement);
		};
		reader.onerror = () => resolve(null);
		reader.readAsDataURL(file);
	});
}
