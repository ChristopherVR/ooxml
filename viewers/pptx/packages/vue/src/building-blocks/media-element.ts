import { classifyMediaType } from 'ooxml-ui/pptx';
import { createEditorId } from 'pptx-viewer-core';
import type { PptxElement } from 'pptx-viewer-core';

import type { CanvasSize } from '../viewer/types';

/** Fixed size (px) of an inserted audio player strip. */
const AUDIO_SIZE = { width: 420, height: 64 };
/** Default and cap (px) of an inserted video, aspect kept when it can be probed. */
const VIDEO_MAX_SIZE = { width: 640, height: 360 };

/** A video data URL's natural size, capped to `VIDEO_MAX_SIZE`. */
function probeVideoSize(dataUrl: string): Promise<{ width: number; height: number }> {
	return new Promise((resolve) => {
		const probe = document.createElement('video');
		probe.preload = 'metadata';
		probe.onloadedmetadata = () => {
			const width = probe.videoWidth || VIDEO_MAX_SIZE.width;
			const height = probe.videoHeight || VIDEO_MAX_SIZE.height;
			if (width > VIDEO_MAX_SIZE.width || height > VIDEO_MAX_SIZE.height) {
				const scale = Math.min(VIDEO_MAX_SIZE.width / width, VIDEO_MAX_SIZE.height / height);
				resolve({ width: Math.round(width * scale), height: Math.round(height * scale) });
				return;
			}
			resolve({ width, height });
		};
		probe.onerror = () => resolve(VIDEO_MAX_SIZE);
		probe.src = dataUrl;
	});
}

/**
 * A centred media element for an audio or video file the host holds (its own picker), sized as
 * Insert > Media sizes it. Resolves null for an unsupported type or an unreadable file.
 */
export function mediaElementFromFile(file: File, canvas: CanvasSize): Promise<PptxElement | null> {
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
			const size = mediaType === 'audio' ? AUDIO_SIZE : await probeVideoSize(dataUrl);
			resolve({
				id: createEditorId('media'),
				type: 'media',
				mediaType,
				mediaMimeType: file.type || undefined,
				mediaData: dataUrl,
				x: Math.max(0, Math.round((canvas.width - size.width) / 2)),
				y: Math.max(0, Math.round((canvas.height - size.height) / 2)),
				width: size.width,
				height: size.height,
			} as unknown as PptxElement);
		};
		reader.onerror = () => resolve(null);
		reader.readAsDataURL(file);
	});
}
