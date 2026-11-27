// @vitest-environment happy-dom
import { afterEach, describe, expect, it, vi } from 'vitest';

import { decodePdfImage, drawableToPdfImage } from '../../export/pdf-image';
import type { PdfImageResource } from 'ooxml-core/pdf/write';
import { videoPicture } from './pdf-vector-media';

vi.mock('../../export/pdf-image', () => ({
	decodePdfImage: vi.fn(),
	drawableToPdfImage: vi.fn(),
}));

const POSTER = 'data:image/png;base64,iVBORw0KGgo=';
const posterImage: PdfImageResource = { width: 900, height: 506, data: new Uint8Array([1]) };
const frameImage: PdfImageResource = { width: 320, height: 180, data: new Uint8Array([2]) };

/** A mounted background clip: a poster, paused at its start, never played. */
function backgroundClip(): HTMLVideoElement {
	const video = document.createElement('video');
	video.setAttribute('poster', POSTER);
	return video;
}

describe('loading the picture a stage video shows', () => {
	afterEach(() => vi.mocked(decodePdfImage).mockReset());
	afterEach(() => vi.mocked(drawableToPdfImage).mockReset());

	it('decodes the poster, not the clip frame, for a background clip that has not played', async () => {
		vi.mocked(decodePdfImage).mockResolvedValue(posterImage);
		vi.mocked(drawableToPdfImage).mockResolvedValue(frameImage);
		expect(await videoPicture(document, backgroundClip())).toBe(posterImage);
		expect(decodePdfImage).toHaveBeenCalledWith(document, POSTER);
		expect(drawableToPdfImage).not.toHaveBeenCalled();
	});

	it('falls back to the current frame when the poster cannot be decoded', async () => {
		vi.mocked(decodePdfImage).mockResolvedValue(undefined);
		vi.mocked(drawableToPdfImage).mockResolvedValue(frameImage);
		expect(await videoPicture(document, backgroundClip())).toBe(frameImage);
	});

	it('draws the current frame of a video without a poster', async () => {
		vi.mocked(drawableToPdfImage).mockResolvedValue(frameImage);
		expect(await videoPicture(document, document.createElement('video'))).toBe(frameImage);
		expect(decodePdfImage).not.toHaveBeenCalled();
	});
});
