/**
 * The inspector's media trim fields (PowerPoint's Trim Audio / Trim Video
 * start and end times), as field descriptors every binding draws with its one
 * generic renderer, bounded at render time by the clip's live duration
 * (`maxOf`).
 *
 * The target is the clip with the duration its preview reports
 * ({@link MediaTrimTarget}); times are shown in seconds. The end is shown as
 * the absolute stop time a user thinks in, while the element keeps
 * `p14:trim/@end`'s distance from the clip's tail (`media-trim-range.ts`); the
 * end field appears once the duration is known, since there is no tail to
 * measure from before. The start cannot pass the end, nor the end the clip.
 * The scrub-bar timeline and Reset Trim stay native.
 *
 * @module render/inspector/media-trim-fields
 */
import type { MediaPptxElement, PptxElement } from 'ooxml-core/pptx';

import { mediaTrimEndAbsoluteMs, mediaTrimEndMsFromAbsoluteMs } from '../media-trim-range';
import type { InspectorSectionDescriptor } from './inspector-field-model';

/** A clip and the duration its preview reports (0 while unknown). */
export interface MediaTrimTarget {
	media: MediaPptxElement;
	durationMs: number;
}

export type MediaTrimSection = InspectorSectionDescriptor<MediaTrimTarget, Partial<PptxElement>>;

const seconds = (ms: number): number => Math.round(ms) / 1000;

/** The clip's stop time, in ms from its start. */
const endMs = (target: MediaTrimTarget): number =>
	mediaTrimEndAbsoluteMs(target.durationMs, target.media.trimEndMs ?? 0);

/** The start and end times (untitled: the media pane heads them). */
export const MEDIA_TRIM_SECTION: MediaTrimSection = {
	id: 'mediaTrim',
	fields: [
		{
			id: 'trimStart',
			kind: 'number',
			labelKey: 'pptx.media.trimStart',
			min: 0,
			step: 0.1,
			unit: 's',
			decimals: 3,
			testId: 'media-trim-start',
			maxOf: (target) => (target.durationMs > 0 ? seconds(endMs(target)) : undefined),
			read: (target) => seconds(target.media.trimStartMs ?? 0),
			write: (_target, value) =>
				({ trimStartMs: Math.round(Number(value) * 1000) }) as Partial<PptxElement>,
		},
		{
			id: 'trimEnd',
			kind: 'number',
			labelKey: 'pptx.media.trimEnd',
			step: 0.1,
			unit: 's',
			decimals: 3,
			testId: 'media-trim-end',
			visibleWhen: (target) => target.durationMs > 0,
			maxOf: (target) => seconds(target.durationMs),
			read: (target) => seconds(endMs(target)),
			write: (target, value) => {
				const startMs = target.media.trimStartMs ?? 0;
				const absolute = Math.max(startMs, Math.round(Number(value) * 1000));
				return {
					trimEndMs: mediaTrimEndMsFromAbsoluteMs(target.durationMs, absolute),
				} as Partial<PptxElement>;
			},
		},
	],
};

/** The trim fields as a section list, what a generic renderer takes. */
export const MEDIA_TRIM_SECTIONS: readonly MediaTrimSection[] = [MEDIA_TRIM_SECTION];
