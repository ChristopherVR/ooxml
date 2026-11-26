/**
 * The inspector's media playback options (PowerPoint's Playback tab: volume,
 * start, play across slides, loop, full screen, hide while not playing, fades
 * and speed), as field descriptors every binding draws with its one generic
 * renderer.
 *
 * The descriptors read a `PptxElement` and write a `Partial<PptxElement>` of
 * the media fields they touch; a non-media element reads defaults and writes
 * nothing. A switched-off flag and a zero fade clear their field instead of
 * writing `false` or `0`, so an untouched clip saves as it loaded. Play across
 * slides starts the clip automatically, as PowerPoint's option does. The
 * transport, the trim timeline (`media-trim-range.ts`) and the bookmarks stay
 * native templates beside them.
 *
 * @module render/inspector/media-playback-fields
 */
import type { MediaPptxElement, PptxElement } from 'ooxml-core/pptx';

import type {
	InspectorFieldDescriptor,
	InspectorFieldOption,
	InspectorSectionDescriptor,
} from './inspector-field-model';

export type MediaPlaybackSection = InspectorSectionDescriptor<PptxElement, Partial<PptxElement>>;
type MediaPlaybackField = InspectorFieldDescriptor<PptxElement, Partial<PptxElement>>;

/** Longest fade the fields offer, in seconds. */
const MAX_FADE_SECONDS = 59;

const START_OPTIONS: readonly InspectorFieldOption[] = [
	{ value: 'onClick', labelKey: 'pptx.media.startOnClick' },
	{ value: 'sequence', labelKey: 'pptx.mediaRibbon.inSequence' },
	{ value: 'auto', labelKey: 'pptx.media.startAutomatically' },
];

const mediaOf = (element: PptxElement): MediaPptxElement | undefined =>
	element.type === 'media' ? element : undefined;

/** The patch for a media element; nothing for any other kind. */
const mediaPatch = (element: PptxElement, patch: Partial<MediaPptxElement>): Partial<PptxElement> =>
	element.type === 'media' ? (patch as Partial<PptxElement>) : {};

type MediaFlag = 'loop' | 'playAcrossSlides' | 'fullScreen' | 'hideWhenNotPlaying';

function flag(
	id: MediaFlag,
	labelKey: string,
	visibleWhen?: (media: MediaPptxElement | undefined) => boolean,
): MediaPlaybackField {
	const field: MediaPlaybackField = {
		id,
		kind: 'toggle',
		labelKey,
		testId: `media-${id}`,
		read: (element) => mediaOf(element)?.[id] === true,
		write: (element, value) =>
			mediaPatch(element, {
				[id]: value === true || undefined,
			}),
	};
	if (visibleWhen) {
		field.visibleWhen = (element) => visibleWhen(mediaOf(element));
	}
	return field;
}

function fade(id: 'fadeInDuration' | 'fadeOutDuration', labelKey: string): MediaPlaybackField {
	return {
		id,
		kind: 'number',
		labelKey,
		min: 0,
		max: MAX_FADE_SECONDS,
		step: 0.25,
		testId: `media-${id}`,
		read: (element) => mediaOf(element)?.[id] ?? 0,
		write: (element, value) => mediaPatch(element, { [id]: Number(value) || undefined }),
	};
}

/** The playback options (untitled: the media pane heads them). */
export const MEDIA_PLAYBACK_SECTION: MediaPlaybackSection = {
	id: 'mediaPlayback',
	fields: [
		{
			id: 'volume',
			kind: 'percent',
			labelKey: 'pptx.media.volume',
			step: 1,
			unit: '%',
			decimals: 0,
			wide: true,
			testId: 'media-volume',
			read: (element) => (mediaOf(element)?.muted ? 0 : (mediaOf(element)?.volume ?? 1)) * 100,
			write: (element, value) =>
				mediaPatch(element, {
					volume: Number(value) / 100,
					...(mediaOf(element)?.muted && { muted: false }),
				}),
		},
		{
			id: 'start',
			kind: 'select',
			labelKey: 'pptx.media.startTrigger',
			options: START_OPTIONS,
			wide: true,
			testId: 'media-start',
			read: (element) =>
				mediaOf(element)?.startMode === 'sequence'
					? 'sequence'
					: mediaOf(element)?.autoPlay
						? 'auto'
						: 'onClick',
			write: (element, value) =>
				mediaPatch(element, {
					startMode: value === 'auto' || value === 'sequence' ? value : 'click',
					autoPlay: value === 'auto',
				}),
		},
		flag(
			'playAcrossSlides',
			'pptx.media.playAcrossSlides',
			(media) => media?.mediaType === 'audio',
		),
		flag('loop', 'pptx.media.loop'),
		flag('fullScreen', 'pptx.media.fullScreen', (media) => media?.mediaType !== 'audio'),
		flag('hideWhenNotPlaying', 'pptx.media.hideWhenNotPlaying'),
		fade('fadeInDuration', 'pptx.media.fadeIn'),
		fade('fadeOutDuration', 'pptx.media.fadeOut'),
		{
			id: 'playbackSpeed',
			kind: 'number',
			labelKey: 'pptx.media.playbackSpeed',
			min: 0.25,
			max: 4,
			step: 0.25,
			wide: true,
			testId: 'media-playbackSpeed',
			read: (element) => mediaOf(element)?.playbackSpeed ?? 1,
			write: (element, value) =>
				mediaPatch(element, { playbackSpeed: Number(value) === 1 ? undefined : Number(value) }),
		},
	],
};

/** The playback options as a section list, what a generic renderer takes. */
export const MEDIA_PLAYBACK_SECTIONS: readonly MediaPlaybackSection[] = [MEDIA_PLAYBACK_SECTION];
