// Cropping a picture the way Word's Crop fields do: the `a:srcRect` insets change and the shown
// size follows, so the picture keeps its scale (Word shrinks the extent as a side is cropped).
import { imageRectFromCrop, type CropInsets } from '../geometry/picture-crop';
import type { InlineImage } from './inline-model';

const SIDES = ['cropLeft', 'cropTop', 'cropRight', 'cropBottom'] as const;

/** A picture's four crop insets as fractions of the whole image (0 when not cropped). */
export function pictureCropOf(image: InlineImage): CropInsets {
	const crop = image.pictureFill?.crop ?? {};
	return {
		cropLeft: crop.cropLeft ?? 0,
		cropTop: crop.cropTop ?? 0,
		cropRight: crop.cropRight ?? 0,
		cropBottom: crop.cropBottom ?? 0,
	};
}

/** The picture's whole (uncropped) size in CSS pixels at its current scale. */
export function uncroppedPictureSize(image: InlineImage): { widthPx: number; heightPx: number } {
	const whole = imageRectFromCrop(
		{ x: 0, y: 0, width: image.widthPx, height: image.heightPx },
		pictureCropOf(image),
	);
	return { widthPx: whole.width, heightPx: whole.height };
}

/** Whether insets leave some of the picture showing on both axes. */
export function isValidPictureCrop(crop: CropInsets): boolean {
	return (
		SIDES.every((side) => Number.isFinite(crop[side])) &&
		crop.cropLeft + crop.cropRight < 1 &&
		crop.cropTop + crop.cropBottom < 1
	);
}

/**
 * A copy of `image` cropped to `crop` (fractions of the whole image; negative values pad). The
 * shown size is rescaled so the picture keeps its scale; a zero inset is dropped, and no inset at
 * all leaves `crop` empty so the writer clears `a:srcRect`. Throws when nothing would show.
 */
export function setPictureCrop(image: InlineImage, crop: CropInsets): InlineImage {
	if (image.vmlFloat) throw new Error('Cropping is unavailable for legacy VML pictures.');
	if (!image.pictureFill) throw new Error('Only a picture can be cropped.');
	if (!isValidPictureCrop(crop))
		throw new RangeError('The crop must leave part of the picture showing.');
	const whole = uncroppedPictureSize(image);
	const next: Partial<CropInsets> = {};
	for (const side of SIDES) {
		const value = Math.round(crop[side] * 100000) / 100000;
		if (value !== 0) next[side] = value;
	}
	const spanX = 1 - (next.cropLeft ?? 0) - (next.cropRight ?? 0);
	const spanY = 1 - (next.cropTop ?? 0) - (next.cropBottom ?? 0);
	return {
		...image,
		widthPx: Math.max(1, whole.widthPx * spanX),
		heightPx: Math.max(1, whole.heightPx * spanY),
		pictureFill: { ...image.pictureFill, crop: next },
	};
}
