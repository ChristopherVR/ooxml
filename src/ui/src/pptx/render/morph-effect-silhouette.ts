/**
 * A morphing shape's silhouette shadow, glow and inner shadow, dissolved the way
 * PowerPoint dissolves them (see `morph-effect-carry`): the outgoing end's
 * effects, riding the box at `1 - p`, and the incoming end's at `p`.
 *
 * Both ends are layers of ONE filter on the arriving shape (`silhouette-morph-
 * registry` hands it the outgoing end, scaled into the arriving frame). Each
 * layer reads the morph's eased progress from the `--pptx-morph-p` custom
 * property, which this module animates on the same container as the other
 * effect keyframes (an `@property` registration makes it interpolate).
 *
 * @module render/morph-effect-silhouette
 */
import type { PptxElement } from 'ooxml-core/pptx';
import { hasShapeProperties } from 'ooxml-core/pptx';

import { MORPH_PROGRESS_VAR } from './silhouette-effects';
import type { SilhouetteEffectParams } from './silhouette-effects';
import {
	getSilhouetteEffectParams,
	resolveGlow,
	resolveInnerShadow,
	resolveOuterShadowLayers,
} from './silhouette-effects';
import { registerSilhouetteMorph } from './silhouette-morph-registry';

/** Registers the property so the keyframes interpolate it instead of flipping it. */
export const MORPH_PROGRESS_PROPERTY = `@property ${MORPH_PROGRESS_VAR} {\n\tsyntax: '<number>';\n\tinherits: true;\n\tinitial-value: 1;\n}`;

/** The effects of an end as a silhouette would paint them, whether or not it qualifies. */
function rawParams(element: PptxElement): SilhouetteEffectParams | undefined {
	if (element.type !== 'shape' || !hasShapeProperties(element) || !element.shapeStyle) {
		return undefined;
	}
	const style = element.shapeStyle;
	const shadows = resolveOuterShadowLayers(style, element.rotation);
	const glow = resolveGlow(style);
	const inner = resolveInnerShadow(style);
	if (shadows.length === 0 && !glow && !inner) {
		return undefined;
	}
	return { shadows, ...(glow ? { glow } : {}), ...(inner ? { inner } : {}) };
}

/** `params` in a frame `kx` x `ky` times the one it was written in. */
function scaleParams(
	params: SilhouetteEffectParams,
	kx: number,
	ky: number,
): SilhouetteEffectParams {
	const mean = Math.sqrt(kx * ky);
	return {
		shadows: params.shadows.map((s) => ({
			...s,
			offsetX: s.offsetX * kx,
			offsetY: s.offsetY * ky,
			blur: s.blur * mean,
		})),
		...(params.glow
			? {
					glow: {
						...params.glow,
						spread: params.glow.spread * mean,
						blur: params.glow.blur * mean,
					},
				}
			: {}),
		...(params.inner
			? {
					inner: {
						...params.inner,
						offsetX: params.inner.offsetX * kx,
						offsetY: params.inner.offsetY * ky,
						blur: params.inner.blur * mean,
					},
				}
			: {}),
	};
}

/**
 * Registers the dissolve of an arriving silhouette shape and returns whether it
 * applies (`underGhost`: the outgoing end's effects stay with the ghost) (the arriving end paints a silhouette and the outgoing one has effects
 * that differ from it once scaled into the arriving frame).
 */
export function registerSilhouetteDissolve(
	from: PptxElement,
	to: PptxElement,
	painter: 'alone' | 'underGhost' = 'alone',
): boolean {
	const own = getSilhouetteEffectParams(to);
	if (!own) {
		return false;
	}
	// A shape arriving from one without effects fades its own in against nothing.
	// Under a ghost the ghost paints the outgoing effects itself, fading with it.
	const outgoing = (painter === 'alone' ? rawParams(from) : undefined) ?? { shadows: [] };
	const out = scaleParams(
		outgoing,
		Math.max(to.width, 1) / Math.max(from.width, 1),
		Math.max(to.height, 1) / Math.max(from.height, 1),
	);
	if (JSON.stringify(out) === JSON.stringify(own)) {
		return false;
	}
	registerSilhouetteMorph(to.id, own, out, { width: to.width, height: to.height });
	return true;
}
