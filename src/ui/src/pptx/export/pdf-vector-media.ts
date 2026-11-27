/**
 * What a `<video>` on the stage shows, for vector PDF export. A browser paints a video's poster
 * until there is a frame to show or playback starts (the HTML "show poster" flag), so a paused
 * video that has never played, such as a slide's looping background clip, shows its poster
 * picture, not the clip's first frame. The export draws what the stage shows.
 */
import { decodePdfImage, drawableToPdfImage } from '../../export/pdf-image';
import type { PdfImageResource } from 'ooxml-core/pdf/write';

/** `HTMLMediaElement.HAVE_CURRENT_DATA`: a frame for the current position is decoded. */
const HAVE_CURRENT_DATA = 2;

type VideoState = Pick<
	HTMLVideoElement,
	'poster' | 'readyState' | 'paused' | 'played' | 'currentTime' | 'getAttribute'
>;

/** The poster the video is showing, or `undefined` when it shows a frame (or has no poster). */
export function shownVideoPoster(video: VideoState): string | undefined {
	if (!video.getAttribute('poster')) return undefined;
	const noFrame = video.readyState < HAVE_CURRENT_DATA;
	const notStarted = video.paused && video.played.length === 0 && video.currentTime === 0;
	return noFrame || notStarted
		? video.poster || video.getAttribute('poster') || undefined
		: undefined;
}

/**
 * Load the picture a video shows: its poster while that is what the stage paints (falling back
 * to the current frame when the poster cannot be decoded), otherwise the current frame.
 */
export async function videoPicture(
	doc: Document,
	video: HTMLVideoElement,
): Promise<PdfImageResource | undefined> {
	const frame = () => drawableToPdfImage(doc, video, video.videoWidth, video.videoHeight);
	const poster = shownVideoPoster(video);
	if (!poster) return frame();
	return (await decodePdfImage(doc, poster)) ?? frame();
}
