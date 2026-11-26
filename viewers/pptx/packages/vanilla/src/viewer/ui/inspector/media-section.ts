import {
	MEDIA_TRIM_SECTIONS,
	mediaTrimResetPatch,
	mediaTrimResetState,
	MEDIA_PLAYBACK_SECTIONS,
	formatMediaBookmarkTime,
	parseMediaBookmarkTime,
} from 'ooxml-ui/pptx';

import type { Translator } from '../../i18n';
import { createInspectorFields } from './inspector-fields';
import { createMediaTrimTimeline } from './media-trim-timeline';
import type { InspectorHandlers, InspectorState } from './types';

export function createMediaSection(
	doc: Document,
	t: Translator,
	section: (label: string) => HTMLElement,
	handlers: InspectorHandlers,
) {
	const el = section(t('pptx.media.title'));
	// The trim fields (`MEDIA_TRIM_SECTIONS`) show the absolute end and are
	// bounded by the clip's live duration, read off the preview when it loads.
	let currentMedia: InspectorState['media'];
	const resolveDurationMs = (): number => {
		const liveSeconds = Number.isFinite(preview.duration) ? preview.duration : 0;
		const metadataSeconds = currentMedia?.metadata?.duration ?? 0;
		return (liveSeconds || metadataSeconds) * 1000;
	};
	const preview = doc.createElement('video');
	preview.controls = true;
	preview.preload = 'metadata';
	preview.className = 'pptxv-media-preview';
	preview.addEventListener('loadedmetadata', () => {
		handlers.setMediaProperties({
			metadata: {
				duration: Number.isFinite(preview.duration) ? preview.duration : undefined,
				videoWidth: preview.videoWidth || undefined,
				videoHeight: preview.videoHeight || undefined,
			},
		});
		paintTimeline();
		paintTrim();
	});
	const poster = doc.createElement('img');
	poster.className = 'pptxv-media-poster';
	poster.alt = t('pptx.media.posterFrame');
	// Volume, start, play across slides, loop, full screen, hide while not
	// playing, fades and speed: the shared descriptors.
	const playback = createInspectorFields(doc, t, {
		sections: MEDIA_PLAYBACK_SECTIONS,
		onPatch: (patch) =>
			handlers.setMediaProperties(patch as Parameters<InspectorHandlers['setMediaProperties']>[0]),
	});
	// Start and end times, bounded by the clip's live duration: the shared descriptors.
	const trim = createInspectorFields(doc, t, {
		sections: MEDIA_TRIM_SECTIONS,
		onPatch: (patch) =>
			handlers.setMediaProperties(patch as Parameters<InspectorHandlers['setMediaProperties']>[0]),
	});
	const paintTrim = (): void => {
		if (currentMedia) {
			trim.update(
				{ media: currentMedia, durationMs: resolveDurationMs() },
				!canEditMedia,
				currentMedia.id,
			);
		}
	};
	const resetTrim = doc.createElement('button');
	resetTrim.type = 'button';
	resetTrim.className = 'pptxv-media-reset-trim';
	resetTrim.textContent = t('pptx.media.resetTrim');
	resetTrim.addEventListener('click', () =>
		handlers.setMediaProperties(
			mediaTrimResetPatch() as Parameters<InspectorHandlers['setMediaProperties']>[0],
		),
	);
	const bookmarks = textArea(doc, t('pptx.media.bookmarks'));
	bookmarks.control.placeholder = formatMediaBookmarkTime(0);
	bookmarks.control.addEventListener('change', () =>
		handlers.setMediaProperties({
			bookmarks: rows(bookmarks.control.value).flatMap((row, index) => {
				const [clock = '', ...labelParts] = row.split('|');
				const label = labelParts.join('|').trim();
				const previous = currentMedia?.bookmarks?.[index];
				const time = parseMediaBookmarkTime(clock, previous?.time);
				return time === undefined
					? []
					: [{ id: previous?.id ?? `bookmark-${index + 1}`, time, label }];
			}),
		}),
	);
	const captions = textArea(doc, t('pptx.media.captions'));
	captions.control.placeholder = 'en | English | captions | captions.vtt';
	captions.control.addEventListener('change', () =>
		handlers.setMediaProperties({
			captionTracks: rows(captions.control.value).map((row, index) => {
				const [language = '', label = '', kind = 'captions', src = ''] = row
					.split('|')
					.map((value) => value.trim());
				return {
					id: `caption-${index + 1}`,
					language,
					label,
					kind: kind as 'subtitles' | 'captions' | 'descriptions',
					src: src || undefined,
				};
			}),
		}),
	);
	const metadata = doc.createElement('output');
	metadata.className = 'pptxv-media-metadata';

	// Scrub bar with draggable trim handles + bookmark ticks (React's
	// `TrimTimeline`). It reads the live `<video>` for duration/playhead, so it
	// only becomes useful once the preview reports metadata; the numeric trim
	// fields above stay as the keyboard-accessible way to set the same values.
	let canEditMedia = false;
	const timeline = createMediaTrimTimeline(doc, {
		onTrimChange: (trimStartMs, trimEndMs) => {
			handlers.setMediaProperties({ trimStartMs, trimEndMs });
		},
		onSeek: (seconds) => {
			if (Number.isFinite(preview.duration)) {
				preview.currentTime = seconds;
			}
		},
	});
	const paintTimeline = (): void => {
		timeline.update({
			duration: Number.isFinite(preview.duration) ? preview.duration : 0,
			trimStartMs: currentMedia?.trimStartMs ?? 0,
			trimEndMs: currentMedia?.trimEndMs ?? 0,
			currentTime: preview.currentTime || 0,
			bookmarks: currentMedia?.bookmarks ?? [],
			canEdit: canEditMedia,
		});
	};
	preview.addEventListener('timeupdate', paintTimeline);

	el.append(
		preview,
		poster,
		timeline.el,
		playback.el,
		trim.el,
		resetTrim,
		bookmarks.label,
		captions.label,
		metadata,
	);
	return {
		el,
		update(state: InspectorState) {
			el.hidden = !state.isMedia;
			const media = state.media;
			currentMedia = media;
			canEditMedia = state.isMedia;
			paintTimeline();
			const source = state.mediaPreviewUrl ?? '';
			if (preview.src !== source) {
				preview.src = source;
			}
			preview.hidden = !source;
			preview.poster = state.mediaPosterUrl ?? '';
			poster.src = state.mediaPosterUrl ?? '';
			poster.hidden = !state.mediaPosterUrl;
			if (media) {
				playback.update(media, !state.isMedia, media.id);
			}
			paintTrim();
			bookmarks.control.value = (media?.bookmarks ?? [])
				.map(({ time, label }) => `${formatMediaBookmarkTime(time)} | ${label}`)
				.join('\n');
			captions.control.value = (media?.captionTracks ?? [])
				.map(({ language, label, kind, src }) => `${language} | ${label} | ${kind} | ${src ?? ''}`)
				.join('\n');
			metadata.textContent = media?.metadata
				? `${media.metadata.duration ?? 0}s, ${media.metadata.videoWidth ?? 0} x ${media.metadata.videoHeight ?? 0}, ${media.metadata.codecInfo ?? media.mediaMimeType ?? ''}`
				: (media?.mediaMimeType ?? '');
			resetTrim.hidden = !state.isMedia || !mediaTrimResetState(media, true).visible;
			bookmarks.control.disabled = !state.isMedia;
			captions.control.disabled = !state.isMedia;
		},
	};
}

function textArea(doc: Document, text: string) {
	const label = doc.createElement('label');
	label.textContent = text;
	const control = doc.createElement('textarea');
	control.rows = 3;
	label.appendChild(control);
	return { label, control };
}

function rows(value: string): string[] {
	return value
		.split(/\r?\n/)
		.map((row) => row.trim())
		.filter(Boolean);
}
