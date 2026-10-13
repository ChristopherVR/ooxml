import type { PptxElement } from 'ooxml-core/pptx';
import type { ElementAnimationState } from './animation-timeline-group';
import { BACKGROUND_ANIMATION_ID_SUFFIX } from './animation-target-id';
import { getShapeEffectAnimationLayerStyle } from './picture-frame';
import type { CssStyleMap } from './element-style-transform';

/** Text/shape paint targets and the existing picture-frame paint path. */
export function getBackgroundAnimationState(
	element: PptxElement,
	states: ReadonlyMap<string, ElementAnimationState> | undefined,
): ElementAnimationState | undefined {
	return element.type === 'text' ||
		element.type === 'shape' ||
		element.type === 'picture' ||
		element.type === 'image'
		? states?.get(`${element.id}${BACKGROUND_ANIMATION_ID_SUFFIX}`)
		: undefined;
}

const PAINT = /^(?:background|border|stroke|fill)|^(?:filter|boxShadow|mixBlendMode|clipPath)$/;

/** Move visual paint off the positioning box so a background target leaves text visible. */
export function getBackgroundAnimationLayers<T extends object>(
	element: PptxElement,
	style: T,
	state: ElementAnimationState | undefined,
): { box: T; paint: CssStyleMap | undefined; targetId: string | undefined } {
	if (!state) return { box: style, paint: undefined, targetId: undefined };
	const box = {},
		paint: CssStyleMap = {};
	for (const [key, value] of Object.entries(style)) {
		if (PAINT.test(key) && (typeof value === 'string' || typeof value === 'number'))
			paint[key] = value;
		else Object.assign(box, { [key]: value });
	}
	return {
		box: box as T,
		paint: getShapeEffectAnimationLayerStyle(element, paint, state),
		targetId: `${element.id}${BACKGROUND_ANIMATION_ID_SUFFIX}`,
	};
}
