/**
 * Which container and codec a video export records, shared by every binding's `MediaRecorder`
 * driver and its download name.
 *
 * Chromium and Firefox record WebM; Safari (macOS, iOS and iPadOS) records only MP4 (H.264). The
 * first candidate `MediaRecorder.isTypeSupported` accepts wins, and when it accepts none the export
 * is reported as unsupported instead of constructing a recorder that throws `NotSupportedError`.
 */
import { translationsEn } from '../i18n/translations-en';
import { WEBM_MIME_CANDIDATES } from './video-plan';

/** Recording MIME types to try, most preferred first: WebM (VP9, VP8), then MP4 (H.264). */
export const VIDEO_MIME_CANDIDATES: readonly string[] = [
	...WEBM_MIME_CANDIDATES,
	'video/mp4;codecs=avc1.42E01E',
	'video/mp4;codecs=avc1',
	'video/mp4',
];

/** The container a recording is saved in. */
export type VideoContainer = 'webm' | 'mp4';

/** The recorder MIME type plus the blob type and file extension that go with it. */
export interface VideoRecordingFormat {
	/** Passed to `new MediaRecorder(stream, { mimeType })`. */
	mimeType: string;
	/** The container, which is also the download's file extension. */
	container: VideoContainer;
	/** The `Blob` type of the assembled recording (`video/webm` or `video/mp4`). */
	blobType: `video/${VideoContainer}`;
}

/** Thrown when the browser can record none of {@link VIDEO_MIME_CANDIDATES}. */
export class VideoExportUnsupportedError extends Error {
	constructor(message: string = translationsEn['pptx.export.errorVideoUnsupported']) {
		super(message);
		this.name = 'VideoExportUnsupportedError';
	}
}

/** The container a recorder MIME type (or blob type) records into. */
export function videoContainerOf(mimeType: string): VideoContainer {
	return /^video\/mp4\b/i.test(mimeType.trim()) ? 'mp4' : 'webm';
}

/** The {@link VideoRecordingFormat} for one recorder MIME type. */
export function videoFormatOf(mimeType: string): VideoRecordingFormat {
	const container = videoContainerOf(mimeType);
	return { mimeType, container, blobType: `video/${container}` };
}

type TypeSupport = (mimeType: string) => boolean;

function browserTypeSupport(): TypeSupport | undefined {
	if (typeof MediaRecorder === 'undefined') return undefined;
	const isTypeSupported = MediaRecorder.isTypeSupported as TypeSupport | undefined;
	return typeof isTypeSupported === 'function'
		? (mimeType) => isTypeSupported.call(MediaRecorder, mimeType)
		: undefined;
}

/**
 * The first candidate the browser can record, or `null` when it can record none. Where
 * `MediaRecorder.isTypeSupported` is missing (no `MediaRecorder` global, as in a test that injects
 * its own recorder), support cannot be asked, so the first candidate is returned and the recorder
 * itself decides.
 */
export function pickVideoRecordingFormat(
	candidates: readonly string[] = VIDEO_MIME_CANDIDATES,
	isTypeSupported: TypeSupport | undefined = browserTypeSupport(),
): VideoRecordingFormat | null {
	const first = candidates[0];
	if (first === undefined) return null;
	if (!isTypeSupported) return videoFormatOf(first);
	for (const mimeType of candidates) {
		try {
			if (isTypeSupported(mimeType)) return videoFormatOf(mimeType);
		} catch {
			// A browser that throws on an unknown type does not support it.
		}
	}
	return null;
}

/** {@link pickVideoRecordingFormat}, throwing {@link VideoExportUnsupportedError} when none is supported. */
export function requireVideoRecordingFormat(
	candidates?: readonly string[],
	isTypeSupported?: TypeSupport,
): VideoRecordingFormat {
	const format =
		isTypeSupported === undefined
			? pickVideoRecordingFormat(candidates)
			: pickVideoRecordingFormat(candidates, isTypeSupported);
	if (!format) throw new VideoExportUnsupportedError();
	return format;
}

/** The download name for a recorded video: `<baseName>.mp4` or `<baseName>.webm` from its type. */
export function videoFileName(
	baseName: string,
	video: Blob | VideoRecordingFormat | string,
): string {
	const type =
		typeof video === 'string' ? video : 'blobType' in video ? video.blobType : video.type;
	return `${baseName}.${videoContainerOf(type)}`;
}
