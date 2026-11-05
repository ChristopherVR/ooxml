/**
 * The picture effects of a new picture: `ImageOptions.opacity`, `brightness`
 * and `contrast` as the `a:blip` effects PowerPoint reads (`a:alphaModFix`
 * and `a:lum`).
 *
 * The effect list is built by the shared DrawingML picture-effect model
 * (`patchPictureEffects`), which owns the ranges and the order of the
 * effects, and is then expressed in the pptx `PptxImageEffects` fields the
 * pptx picture writer saves.
 *
 * @module sdk/picture-effects-input
 */

import { patchPictureEffects } from '../../../../drawingml/picture-effects';
import type { PictureEffectPatch } from '../../../../drawingml/picture-effects';
import type { PptxImageEffects } from '../../types/image';
import type { ImageOptions } from './types';

/**
 * The `imageEffects` of a new picture, or `undefined` when no option asks
 * for one (full opacity and no brightness / contrast change write nothing).
 *
 * @throws RangeError for an opacity outside 0..1 or a brightness / contrast
 *   outside -100..100, or a value that is not finite.
 */
export function pictureEffectsFromInput(
	options: Pick<ImageOptions, 'opacity' | 'brightness' | 'contrast'> | undefined,
): PptxImageEffects | undefined {
	const { opacity, brightness, contrast } = options ?? {};
	if (opacity !== undefined && !(opacity >= 0 && opacity <= 1)) {
		throw new RangeError(`ImageOptions.opacity must be between 0 and 1, got ${opacity}`);
	}
	const patch: PictureEffectPatch = {};
	if (opacity !== undefined && opacity < 1) {
		patch.alpha = opacity;
	}
	if (brightness !== undefined) {
		patch.brightness = brightness / 100;
	}
	if (contrast !== undefined) {
		patch.contrast = contrast / 100;
	}
	if (Object.keys(patch).length === 0) {
		return undefined;
	}
	let patched: ReturnType<typeof patchPictureEffects>;
	try {
		patched = patchPictureEffects([], patch);
	} catch (error) {
		throw new RangeError(
			`ImageOptions.brightness and contrast must be finite percentages from -100 to 100 (${String(error)})`,
		);
	}
	let effects: PptxImageEffects | undefined;
	for (const effect of patched) {
		effects ??= {};
		if (effect.kind === 'alphaModFix') {
			effects.alphaModFix = effect.amount * 100;
		} else if (effect.kind === 'luminance') {
			effects.lum = {
				...(effect.brightness !== undefined ? { bright: effect.brightness * 100 } : {}),
				...(effect.contrast !== undefined ? { contrast: effect.contrast * 100 } : {}),
			};
		}
	}
	return effects;
}
