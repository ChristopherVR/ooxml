/**
 * The inspector's picture fields (PowerPoint's Format Picture > Picture):
 * the adjustments (brightness, contrast, saturation, transparency, black and
 * white threshold, grayscale) and the four crop insets, as field descriptors
 * every binding draws with its one generic renderer.
 *
 * The descriptors read a `PptxElement` and write a `Partial<PptxElement>`
 * (an `imageEffects` merged over the picture's own, or crop fractions), and a
 * non-picture reads neutral values and writes nothing. The compound picture
 * controls (artistic effects, duotone, colour wash, recolour, crop to shape,
 * replace and reset) stay native templates beside them.
 *
 * Transparency is shown the way PowerPoint shows it, inverted from the stored
 * opacity (`a:alphaModFix/@amt`, 100 = opaque). Zero transparency and a zero
 * threshold clear their effect instead of writing it: a threshold of 0 would
 * turn every pixel white. Crop insets are percentages of their edge, bounded
 * by the shared crop clamp (90%).
 *
 * @module render/inspector/picture-fields
 */
import type { PptxElement, PptxImageEffects } from 'ooxml-core/pptx';
import { isImageLikeElement } from 'ooxml-core/pptx';

import { imageAdjustmentsPatch, imageCropPatch, imageCropStateOf } from '../image-adjustments';
import type { ImageCropState } from '../image-adjustments';
import type { InspectorFieldDescriptor, InspectorSectionDescriptor } from './inspector-field-model';

export type PictureSection = InspectorSectionDescriptor<PptxElement, Partial<PptxElement>>;
type PictureField = InspectorFieldDescriptor<PptxElement, Partial<PptxElement>>;

/** Largest crop inset the fields offer, in percent (the shared crop clamp). */
export const MAX_PICTURE_CROP_PERCENT = 90;

const effectsOf = (element: PptxElement): PptxImageEffects | undefined =>
	isImageLikeElement(element) ? element.imageEffects : undefined;

/** Merge `changes` into the picture's `imageEffects`; nothing for a non-picture. */
function effectsPatch(
	element: PptxElement,
	changes: Partial<PptxImageEffects>,
): Partial<PptxElement> {
	if (!isImageLikeElement(element)) {
		return {};
	}
	return { imageEffects: { ...element.imageEffects, ...changes } } as Partial<PptxElement>;
}

function adjustment(id: 'brightness' | 'contrast' | 'saturation', labelKey: string): PictureField {
	return {
		id,
		kind: 'percent',
		labelKey,
		min: -100,
		max: 100,
		step: 1,
		unit: '%',
		decimals: 0,
		wide: true,
		testId: `picture-${id}`,
		read: (element) => effectsOf(element)?.[id] ?? 0,
		write: (element, value) => imageAdjustmentsPatch(element, { [id]: Number(value) }),
	};
}

/** The picture adjustments (untitled: each pane heads its picture card). */
export const PICTURE_ADJUSTMENT_SECTION: PictureSection = {
	id: 'pictureAdjustments',
	fields: [
		adjustment('brightness', 'pptx.imageAdjustments.brightness'),
		adjustment('contrast', 'pptx.imageAdjustments.contrast'),
		adjustment('saturation', 'pptx.image.saturation'),
		{
			id: 'transparency',
			kind: 'percent',
			labelKey: 'pptx.imageAdjustments.transparency',
			step: 1,
			unit: '%',
			decimals: 0,
			wide: true,
			testId: 'picture-transparency',
			read: (element) => 100 - (effectsOf(element)?.alphaModFix ?? 100),
			write: (element, value) => {
				const transparency = Number(value);
				return effectsPatch(element, {
					alphaModFix: transparency > 0 ? 100 - transparency : undefined,
				});
			},
		},
		{
			id: 'biLevel',
			kind: 'percent',
			labelKey: 'pptx.imageAdjustments.biLevelThreshold',
			step: 1,
			unit: '%',
			decimals: 0,
			wide: true,
			testId: 'picture-bilevel',
			read: (element) => effectsOf(element)?.biLevel ?? 0,
			write: (element, value) => {
				const threshold = Number(value);
				return effectsPatch(element, { biLevel: threshold > 0 ? threshold : undefined });
			},
		},
		{
			id: 'grayscale',
			kind: 'toggle',
			labelKey: 'pptx.image.grayscale',
			testId: 'picture-grayscale',
			read: (element) => effectsOf(element)?.grayscale === true,
			write: (element, value) => effectsPatch(element, { grayscale: value === true || undefined }),
		},
	],
};

const CROP_EDGES: readonly [keyof ImageCropState, string][] = [
	['cropLeft', 'pptx.image.cropLeft'],
	['cropTop', 'pptx.image.cropTop'],
	['cropRight', 'pptx.image.cropRight'],
	['cropBottom', 'pptx.image.cropBottom'],
];

/**
 * The four crop insets (untitled; a binding draws its Reset Crop and crop to
 * shape controls after them, and disables the section while the picture
 * cannot be cropped, `cropResetState`).
 */
export const PICTURE_CROP_SECTION: PictureSection = {
	id: 'pictureCrop',
	fields: CROP_EDGES.map(([edge, labelKey]): PictureField => ({
		id: edge,
		kind: 'percent',
		labelKey,
		min: 0,
		max: MAX_PICTURE_CROP_PERCENT,
		step: 1,
		unit: '%',
		decimals: 0,
		testId: `picture-${edge}`,
		read: (element) => imageCropStateOf(element)[edge] * 100,
		write: (element, value) => imageCropPatch(element, { [edge]: Number(value) / 100 }),
	})),
};

/** The adjustments as a section list, what a generic renderer takes. */
export const PICTURE_ADJUSTMENT_SECTIONS: readonly PictureSection[] = [PICTURE_ADJUSTMENT_SECTION];
/** The crop insets as a section list. */
export const PICTURE_CROP_SECTIONS: readonly PictureSection[] = [PICTURE_CROP_SECTION];
