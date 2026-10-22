import type { MorphTransitionPlan } from 'ooxml-ui/pptx';
import {
	MORPH_CROSSFADE_GROUP_STYLE,
	MORPH_CROSSFADE_HALF_BLEND_MODE,
	visibleTemplateElements as filterVisibleTemplateElements,
} from 'ooxml-ui/pptx';
import type { PptxElement, PptxSlide } from 'pptx-viewer-core';

import type { StyleMap } from './element-style';

/**
 * Pure morph-transition helpers for `PresentationTransitionOverlayComponent`,
 * split out to keep that file under the project's per-file LOC budget. This
 * package renders no component under test (see
 * `presentation-transition-overlay.component.test.ts`), so every function
 * here is exported and unit-tested directly rather than through a fixture.
 */

/**
 * `slide` with `templateElements` merged ahead of its own, when any are
 * visible. Shared by every consumer that paints a slide "as authored" (the
 * classic incoming layer and the outgoing layer).
 */
export function withTemplateElements(
	slide: PptxSlide,
	templateElements: readonly PptxElement[],
): PptxSlide {
	const visible = filterVisibleTemplateElements(slide, templateElements);
	if (visible.length === 0) {
		return slide;
	}
	return { ...slide, elements: [...visible, ...slide.elements] };
}

/**
 * The slide the incoming (arriving) layer paints for a CLASSIC transition, or
 * `undefined` when no such layer may exist.
 *
 * Without this layer the overlay painted only the outgoing slide: a wipe
 * (whose outgoing half is `none`) sat opaque for the whole duration and the
 * arriving slide - only ever the live stage beneath - popped in the instant
 * the overlay tore down ("takes the time, then instantly replaced"). Types
 * whose incoming half is `none` (the uncover family) deliberately reveal the
 * live stage and get no layer, and a morph paints its own halves.
 */
export function classicIncomingLayerSlide(
	isMorph: boolean,
	incomingAnimation: string,
	incomingSlide: PptxSlide | undefined,
	templateElements: readonly PptxElement[],
): PptxSlide | undefined {
	if (isMorph || incomingAnimation === 'none' || !incomingSlide) {
		return undefined;
	}
	return withTemplateElements(incomingSlide, templateElements);
}

/** The z-index of the morph's first overlay layer, as the ghost layer had. */
const MORPH_LAYER_Z_BASE = 40;

/** One run of the morph overlay, as the slide that paints it. */
export interface MorphOverlayLayerSlide {
	key: string;
	side: 'outgoing' | 'incoming';
	style: StyleMap;
	slide: PptxSlide;
}

/**
 * The ghosts and lifted arrivals of a morph, one stacked layer per run of the
 * plan's `overlayLayers`, back to front.
 *
 * A ghost run renders against the outgoing slide inside a
 * `data-pptx-morph-outgoing` scope. A lifted run holds shapes arriving inside
 * one that persists, which the live stage draws UNDER this overlay where the
 * persisting shape's opaque ghost would hide them (issue #146); it renders
 * against the incoming slide inside a `data-pptx-morph-lifted` scope, and the
 * plan holds their stage copy invisible. One shared order, rather than every
 * ghost under every lifted shape, keeps a full-slide picture fading in from
 * veiling the title travelling above it (issue #141).
 */
export function morphOverlayLayerSlides(
	plan: MorphTransitionPlan | undefined,
	outgoingSlide: PptxSlide | undefined,
	incomingSlide: PptxSlide | undefined,
): MorphOverlayLayerSlide[] {
	if (!plan || !outgoingSlide) {
		return [];
	}
	return plan.overlayLayers.flatMap((layer, index) => {
		const slide = layer.side === 'outgoing' ? outgoingSlide : incomingSlide;
		if (!slide) {
			return [];
		}
		return [
			{
				key: `${layer.side}-${layer.elements[0]?.id ?? index}`,
				side: layer.side,
				style: { 'z-index': String(MORPH_LAYER_Z_BASE + index) },
				slide: { ...slide, elements: [...layer.elements] },
			},
		];
	});
}

/** One cross-dissolving pair, as the two single-element slides that paint it. */
export interface MorphCrossfadeGroupSlides {
	key: string;
	style: StyleMap;
	/**
	 * The departing half's layer style, carrying the dissolve itself.
	 *
	 * A pair dissolving in place never moves, and an animation on the small
	 * element box gives it a compositing layer whose raster snaps to whole device
	 * pixels - the wording is then painted a fraction of a pixel off the live
	 * stage and twitches as the overlay comes and goes (issue #161).
	 */
	outgoingStyle: StyleMap;
	/** The arriving half's layer style. @see MorphCrossfadeGroupSlides.outgoingStyle */
	incomingStyle: StyleMap;
	outgoing: PptxSlide;
	incoming: PptxSlide;
}

/**
 * The pairs the overlay paints BOTH halves of, as one isolated group each.
 *
 * Stacking the halves composites them source-over, which leaves the ink they
 * share at 0.75 of full strength halfway through instead of summing it, biting
 * chunks out of glyphs that cross during a text dissolve. PowerPoint's own
 * render holds the two blend coefficients at a sum of 1.0 for every frame
 * (issue #161), which `isolation: isolate` plus `mix-blend-mode: plus-lighter`
 * on the two halves reproduces.
 */
export function morphCrossfadeGroupSlides(
	plan: MorphTransitionPlan | undefined,
	outgoingSlide: PptxSlide | undefined,
	incomingSlide: PptxSlide | undefined,
): MorphCrossfadeGroupSlides[] {
	if (!plan || !outgoingSlide || !incomingSlide) {
		return [];
	}
	const half = (animation: string | undefined): StyleMap => ({
		'mix-blend-mode': MORPH_CROSSFADE_HALF_BLEND_MODE,
		...(animation === undefined ? {} : { animation }),
	});
	return plan.crossfadeGroups.map((group, index) => ({
		key: group.incoming.id,
		style: {
			...MORPH_CROSSFADE_GROUP_STYLE,
			// `isolation` makes the group a stacking context, so it carries a
			// z-index of its own to stay above every overlay layer.
			'z-index': String(MORPH_LAYER_Z_BASE + plan.overlayLayers.length + index),
		},
		outgoingStyle: half(group.outgoingAnimation),
		incomingStyle: half(group.incomingAnimation),
		outgoing: { ...outgoingSlide, elements: [group.outgoing] },
		incoming: { ...incomingSlide, elements: [group.incoming] },
	}));
}
