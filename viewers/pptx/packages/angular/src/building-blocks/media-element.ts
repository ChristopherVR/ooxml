import { classifyMediaType } from 'ooxml-ui/pptx';
import type { CanvasSize } from 'ooxml-ui/pptx';
import type { PptxElement } from 'pptx-viewer-core';

import { readAsDataUrl } from '../viewer/ribbon-insert-file-picker';

const AUDIO_SIZE = { width: 420, height: 64 };
const VIDEO_MAX = { width: 640, height: 360 };

function probeVideo(dataUrl: string): Promise<{ width: number; height: number }> {
	return new Promise((resolve) => {
		const video = document.createElement('video');
		video.preload = 'metadata';
		video.onloadedmetadata = () => {
			const w = video.videoWidth || VIDEO_MAX.width;
			const h = video.videoHeight || VIDEO_MAX.height;
			const scale = Math.min(1, VIDEO_MAX.width / w, VIDEO_MAX.height / h);
			resolve({ width: Math.round(w * scale), height: Math.round(h * scale) });
		};
		video.onerror = () => resolve(VIDEO_MAX);
		video.src = dataUrl;
	});
}

/** A centred media element for an audio or video file; null when it is not media or unreadable. */
export async function mediaElementFromFile(
	file: File,
	canvas: CanvasSize,
): Promise<PptxElement | null> {
	const mediaType = classifyMediaType(file.type);
	if (!mediaType) {
		return null;
	}
	const dataUrl = await readAsDataUrl(file);
	if (!dataUrl) {
		return null;
	}
	const size = mediaType === 'audio' ? AUDIO_SIZE : await probeVideo(dataUrl);
	return {
		type: 'media',
		id: '',
		name: file.name,
		mediaType,
		mediaMimeType: file.type || undefined,
		mediaData: dataUrl,
		x: Math.round((canvas.width - size.width) / 2),
		y: Math.round((canvas.height - size.height) / 2),
		...size,
	} as PptxElement;
}
