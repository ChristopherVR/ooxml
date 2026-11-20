import { afterEach, describe, expect, it, vi } from 'vitest';
import {
	pickVideoRecordingFormat,
	requireVideoRecordingFormat,
	VIDEO_MIME_CANDIDATES,
	VideoExportUnsupportedError,
	videoFileName,
} from './video-format';

afterEach(() => {
	vi.unstubAllGlobals();
});

describe('video recording format', () => {
	it('prefers WebM where the browser records it', () => {
		const format = pickVideoRecordingFormat(VIDEO_MIME_CANDIDATES, () => true);
		expect(format).toEqual({
			mimeType: 'video/webm;codecs=vp9',
			container: 'webm',
			blobType: 'video/webm',
		});
	});
	it('chooses MP4 on Safari, which records no WebM', () => {
		const format = pickVideoRecordingFormat(VIDEO_MIME_CANDIDATES, (mime) =>
			mime.startsWith('video/mp4'),
		);
		expect(format?.mimeType).toBe('video/mp4;codecs=avc1.42E01E');
		expect(format?.blobType).toBe('video/mp4');
		expect(videoFileName('Deck', format!)).toBe('Deck.mp4');
	});
	it('reads support from MediaRecorder.isTypeSupported', () => {
		const isTypeSupported = vi.fn((mime: string) => mime === 'video/mp4');
		vi.stubGlobal('MediaRecorder', { isTypeSupported });
		expect(pickVideoRecordingFormat()?.mimeType).toBe('video/mp4');
		expect(isTypeSupported).toHaveBeenCalledWith('video/webm;codecs=vp9');
	});
	it('reports unsupported instead of falling back to a type the browser rejects', () => {
		expect(pickVideoRecordingFormat(VIDEO_MIME_CANDIDATES, () => false)).toBeNull();
		expect(() => requireVideoRecordingFormat(VIDEO_MIME_CANDIDATES, () => false)).toThrow(
			VideoExportUnsupportedError,
		);
		const throwing = () => {
			throw new Error('unknown type');
		};
		expect(pickVideoRecordingFormat(['video/x'], throwing)).toBeNull();
	});
	it('lets an injected recorder decide when isTypeSupported is missing', () => {
		vi.stubGlobal('MediaRecorder', undefined);
		expect(pickVideoRecordingFormat()?.mimeType).toBe('video/webm;codecs=vp9');
	});
	it('names the file from a blob type', () => {
		expect(videoFileName('Deck', new Blob([], { type: 'video/webm' }))).toBe('Deck.webm');
		expect(videoFileName('Deck', 'video/mp4;codecs=avc1')).toBe('Deck.mp4');
	});
});
