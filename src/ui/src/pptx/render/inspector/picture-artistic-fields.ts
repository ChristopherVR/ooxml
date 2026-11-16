/**
 * The inspector's artistic effect picker (PowerPoint's Picture Format >
 * Artistic Effects gallery), as a `gallery` field every binding draws with its
 * one generic renderer: one swatch per effect, each previewing the picture
 * through the effect's CSS filter (`ARTISTIC_EFFECTS`).
 *
 * @module render/inspector/picture-artistic-fields
 */
import type { PptxElement } from 'ooxml-core/pptx';
import { isImageLikeElement } from 'ooxml-core/pptx';

import { ARTISTIC_EFFECTS } from '../image-artistic-presets';
import type { InspectorFieldOption, InspectorSectionDescriptor } from './inspector-field-model';

export type PictureArtisticSection = InspectorSectionDescriptor<PptxElement, Partial<PptxElement>>;

const ARTISTIC_OPTIONS: readonly InspectorFieldOption[] = ARTISTIC_EFFECTS.map(
	([value, labelKey, filter]) => {
		const option: InspectorFieldOption = { value, labelKey };
		if (filter) {
			option.previewFilter = filter;
		}
		return option;
	},
);

/** The artistic effect gallery (titled: the gallery has no other heading). */
export const PICTURE_ARTISTIC_SECTION: PictureArtisticSection = {
	id: 'pictureArtistic',
	fields: [
		{
			id: 'artisticEffect',
			kind: 'gallery',
			labelKey: 'pptx.image.artisticEffects',
			options: ARTISTIC_OPTIONS,
			wide: true,
			testId: 'picture-artistic',
			previewImageOf: (element) =>
				isImageLikeElement(element) ? (element.imageData ?? undefined) : undefined,
			read: (element) =>
				(isImageLikeElement(element) ? element.imageEffects?.artisticEffect : undefined) ?? 'none',
			write: (element, value) => {
				if (!isImageLikeElement(element)) {
					return {};
				}
				const effect = String(value);
				return {
					imageEffects: {
						...element.imageEffects,
						artisticEffect: effect === 'none' ? undefined : effect,
						// The grayscale effect replaces the picture's own grayscale toggle.
						...(effect === 'grayscale' ? { grayscale: undefined } : {}),
					},
				} as Partial<PptxElement>;
			},
		},
	],
};

/** The gallery as a section list, what a generic renderer takes. */
export const PICTURE_ARTISTIC_SECTIONS: readonly PictureArtisticSection[] = [
	PICTURE_ARTISTIC_SECTION,
];
