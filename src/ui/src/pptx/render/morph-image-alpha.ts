/**
 * Picture transparency (`a:blip/a:alphaModFix`) across a morph.
 *
 * A common Morph authoring trick fades a picture in or out by putting the SAME
 * picture on both slides and changing only its Transparency: `alphaModFix
 * amt="0"` on one slide, no `amt` (fully opaque) on the next. PowerPoint
 * tweens that transparency over the morph. The engine compared pictures by
 * geometry and image path only, so the pair read as INERT: its ghost was
 * painted statically at the outgoing transparency, the incoming half was held
 * invisible beneath it, and the picture changed in one frame when the overlay
 * came down (issue #141 - a full-slide backdrop popping in at the end of the
 * morph).
 *
 * The transparency is painted as CSS `opacity` on the picture's `<img>` (see
 * `getImageEffectsOpacity`), never on its container, so the tween rides the
 * same `<img>` channel as a crop change (`morph-image-crop`) and composes with
 * whatever the container's own journey does.
 *
 * @module render/morph-image-alpha
 */
import type { PptxElement } from 'ooxml-core/pptx';

import { getImageEffectsOpacity } from './image-effects';
import type { MorphAnimationStyle, MorphPair } from './morph-types';
import { MORPH_EASING } from './morph-types';

/** Below this, two transparencies are the same to the eye (well under 1/255). */
const ALPHA_EPSILON = 0.001;

/** The `<img>` opacity a picture is painted at; 1 when it sets no transparency. */
function imageAlpha(element: PptxElement): number {
	return getImageEffectsOpacity(element) ?? 1;
}

/**
 * Whether a matched pair's picture transparency differs between the slides.
 * Only pictures carry `alphaModFix`, so any other pair compares equal.
 */
export function morphImageAlphaChanged(fromElement: PptxElement, toElement: PptxElement): boolean {
	return Math.abs(imageAlpha(fromElement) - imageAlpha(toElement)) > ALPHA_EPSILON;
}

/** One `<img>` opacity track from the outgoing transparency to the incoming one. */
function alphaTrack(
	index: number,
	suffix: string,
	elementId: string,
	fromElement: PptxElement,
	toElement: PptxElement,
	durationMs: number,
): MorphAnimationStyle {
	const safeName = `pptx-morph-alpha-${suffix}${index}-${elementId.replace(/[^a-zA-Z0-9]/gu, '')}`;
	return {
		elementId,
		target: 'image',
		// The same curve as a matched pair's element-level opacity change, which
		// rides the journey keyframes.
		animation: `${safeName} ${durationMs}ms ${MORPH_EASING} forwards`,
		keyframes: `
@keyframes ${safeName} {
\tfrom {
\t\topacity: ${imageAlpha(fromElement)};
\t}
\tto {
\t\topacity: ${imageAlpha(toElement)};
\t}
}`,
	};
}

/**
 * The INCOMING half of every pair whose picture transparency changed, keyed by
 * the incoming element id and targeted at its `<img>`: it is rendered at its
 * final transparency and started at the outgoing one's.
 *
 * @param pairs - Matched pairs.
 * @param durationMs - Animation duration in milliseconds.
 * @returns One descriptor per pair whose transparency actually changed.
 */
export function generateImageAlphaMorphAnimations(
	pairs: MorphPair[],
	durationMs: number,
): MorphAnimationStyle[] {
	const animations: MorphAnimationStyle[] = [];
	for (let index = 0; index < pairs.length; index++) {
		const { fromElement, toElement } = pairs[index];
		if (morphImageAlphaChanged(fromElement, toElement)) {
			animations.push(alphaTrack(index, '', toElement.id, fromElement, toElement, durationMs));
		}
	}
	return animations;
}

/**
 * The OUTGOING half: the same tween on a ghost the overlay is painting, so a
 * ghost travelling above its counterpart ends at the transparency the live
 * element is left at and the overlay comes down without a seam.
 *
 * @param pairs - Matched pairs.
 * @param durationMs - Animation duration in milliseconds.
 * @param ghostIds - Outgoing ids the overlay will paint; defaults to "all".
 * @returns One descriptor per painted ghost whose transparency changed.
 */
export function generateImageAlphaGhostAnimations(
	pairs: MorphPair[],
	durationMs: number,
	ghostIds?: ReadonlySet<string>,
): MorphAnimationStyle[] {
	const animations: MorphAnimationStyle[] = [];
	for (let index = 0; index < pairs.length; index++) {
		const { fromElement, toElement } = pairs[index];
		if (ghostIds && !ghostIds.has(fromElement.id)) {
			continue;
		}
		if (morphImageAlphaChanged(fromElement, toElement)) {
			animations.push(
				alphaTrack(index, 'ghost-', fromElement.id, fromElement, toElement, durationMs),
			);
		}
	}
	return animations;
}
