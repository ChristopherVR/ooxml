/**
 * The full shape Effects section list, in Format Shape > Effects order.
 *
 * @module render/inspector/shape-effect-sections
 */
import { GLOW, INNER_SHADOW, OUTER_SHADOW, SOFT_EDGE } from './shape-effect-fields';
import type { ShapeEffectSection } from './shape-effect-fields';
import { BEVEL, BLUR, REFLECTION } from './shape-effect-fields-more';

/**
 * Every shape effect section. Bindings draw
 * the whole list with their generic field renderer.
 */
export const SHAPE_EFFECT_SECTIONS: readonly ShapeEffectSection[] = [
	OUTER_SHADOW,
	INNER_SHADOW,
	GLOW,
	SOFT_EDGE,
	REFLECTION,
	BLUR,
	BEVEL,
];
