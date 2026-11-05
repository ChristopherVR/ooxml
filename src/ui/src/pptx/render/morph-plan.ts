/**
 * Turn two slides into everything a binding needs to actually PLAY a morph
 * transition.
 *
 * `morph-animation` already produces the CSS, but as one flat list keyed by a
 * mix of outgoing- and incoming-slide element ids, which is not directly
 * usable by a view layer. This module partitions it the way a renderer
 * consumes it and states the contract the bindings share, so React / Vue /
 * Angular / Svelte / Vanilla do not each re-derive it.
 *
 * ## How a morph is composed
 *
 * PowerPoint's Morph is a FLIP animation, and the generated keyframes follow
 * that model:
 *
 * - A **matched pair** animates on the INCOMING element. It is rendered at its
 *   final geometry and the keyframes start it at the outgoing element's
 *   offset/scale/rotation, so it appears to glide into place. Hence
 *   {@link MorphTransitionPlan.incomingAnimations} is keyed by incoming id.
 *   When the pair's APPEARANCE also changed (fill, outline, picture, text) the
 *   incoming half fades in over a ghost of the outgoing half travelling the
 *   same path, so the two dissolve into each other instead of cutting.
 * - An element only on the **outgoing** slide fades out in place. It is not in
 *   the incoming slide at all, so the binding has to keep painting it for the
 *   duration: {@link MorphTransitionPlan.outgoingElements} is exactly that set,
 *   with its animations in {@link MorphTransitionPlan.outgoingAnimations}.
 * - An element only on the **incoming** slide fades in; also keyed by incoming
 *   id in `incomingAnimations`.
 *
 * A binding therefore: injects {@link MorphTransitionPlan.keyframesCss} once,
 * renders `outgoingElements` in an overlay above the slide, applies
 * `outgoingAnimations` to them and `incomingAnimations` to the real slide's
 * elements, then tears the overlay down after `durationMs`. It must also render
 * {@link MorphTransitionPlan.crossfadeGroups}, the pairs whose two halves are
 * both painted here and have to be blended additively to dissolve the way
 * PowerPoint does; their elements appear in no other list.
 *
 * @module render/morph-plan
 */
import type { PptxElement, PptxSlide } from 'ooxml-core/pptx';

import { generateFullMorphTransition, resolveMorphGhostIds } from './morph-animation';
import { flattenMorphElements } from './morph-flatten';
import { matchMorphElementsFull } from './morph-matching';
import { partitionMorphAnimations } from './morph-plan-channels';
import { planMorphTextCarry } from './morph-plan-carry';
import { resolveMorphLiftedIds } from './morph-plan-lift';
import { assembleMorphOverlay, settleMorphLifts } from './morph-plan-overlay';
import type { MorphPlanOptions, MorphTransitionPlan } from './morph-plan-types';
import type { MorphMode } from './morph-types';

export {
	buildMorphAnimationRules,
	buildMorphOutlineRules,
	buildMorphScopedCss,
} from './morph-plan-css';
export type { MorphPlanOptions, MorphTransitionPlan } from './morph-plan-types';

/** Map a parsed `<p159:morph @option>` onto the engine's granularity mode. */
export function morphOptionToMode(option: string | undefined): MorphMode {
	if (option === 'byWord') {
		return 'word';
	}
	if (option === 'byChar') {
		return 'character';
	}
	return 'object';
}

/**
 * Build the render plan for morphing `fromSlide` into `toSlide`.
 *
 * Returns `undefined` when there is nothing to morph (no matched pairs and
 * nothing to fade), so a caller can fall back to a plain transition.
 */
export function buildMorphTransitionPlan(
	fromSlide: PptxSlide | undefined,
	toSlide: PptxSlide | undefined,
	durationMs: number,
	mode: MorphMode = 'object',
	options: MorphPlanOptions = {},
): MorphTransitionPlan | undefined {
	if (!fromSlide || !toSlide) {
		return undefined;
	}
	const match = matchMorphElementsFull(fromSlide, toSlide);
	if (
		match.pairs.length === 0 &&
		match.unmatchedFrom.length === 0 &&
		match.unmatchedTo.length === 0
	) {
		return undefined;
	}

	const animations = generateFullMorphTransition(
		fromSlide,
		toSlide,
		durationMs,
		mode,
		options.zIndexBase ?? 0,
		options.measureText,
		options.slideSize,
	);

	// The overlay paints the outgoing slide as a moving copy of itself, in the
	// slide's own document order so its z-stacking is preserved:
	//
	//  - shapes with no counterpart fade out in place;
	//  - a matched pair's outgoing half glides onto its counterpart, dissolving
	//    into it when the appearance changed and simply landing on it when only
	//    the geometry did.
	//
	// Painting the unchanged halves too is what keeps a full-slide background
	// from hiding them: the overlay is one flat layer above the live stage, so
	// anything left out of it is invisible behind a crossfading backdrop until
	// that backdrop has faded. Before this, a near-duplicate slide pair (the
	// usual Morph authoring pattern) cut straight to its final look on frame 1
	// and looked like no transition at all (issue #131).
	//
	// Element ids embed their slide path, so the two id spaces do not overlap,
	// but partitioning on this set (rather than on id shape) keeps that an
	// implementation detail of core rather than an assumption here.
	//
	// The list is FLATTENED the same way the matcher flattens it (see
	// `morph-flatten`), against the SAME counterpart, so the two agree on which
	// groups were decomposed: the animations are keyed by the decomposed
	// children's ids. Painting the undecomposed group here instead would paint
	// the children twice over - once inside the group, once as their own ghosts
	// - and leave the group itself without an animation.
	const flattenedOutgoing = flattenMorphElements(fromSlide.elements, toSlide.elements);
	const outgoingIds = new Set(flattenedOutgoing.map((element) => element.id));
	const flattenedIncoming = flattenMorphElements(toSlide.elements, fromSlide.elements);
	const ghostIds = resolveMorphGhostIds(flattenedOutgoing, match.pairs, flattenedIncoming);

	const channels = partitionMorphAnimations(animations, outgoingIds);
	const {
		keyframes,
		incomingAnimations,
		outgoingAnimations,
		outlineAnimations,
		outlineStrandAnimations,
		customGeometryAnimations,
	} = channels;

	// Which outgoing shapes the overlay paints is `resolveMorphGhostIds`' call
	// and nothing else's: a shape whose ghost it drops draws the same thing on
	// the live stage along the same path, and painting it again in the overlay
	// would only hide what is arriving beneath it (issue #144 - the detail
	// slide's callouts never appeared until the overlay came down).
	//
	// Painted is NOT the same question as animated. A ghost standing in for an
	// inert pair is deliberately given no animation, because a running one would
	// put it on its own compositing layer and shift its raster by up to a pixel
	// for the duration (issue #161); it still has to be painted, so this asks
	// the ghost set directly rather than reading it off the animation map.
	const { offstageRestyles, pieces, bodies } = planMorphTextCarry({
		mode,
		durationMs,
		measureText: options.measureText,
		pairs: match.pairs,
		ghostIds,
		flattenedOutgoing,
		flattenedIncoming,
		animated: (id) => incomingAnimations.has(id),
		outlineAnimations,
		customGeometryAnimations,
		outlineStrandAnimations,
	});
	const piecePairs = new Set(pieces.pairs);
	const pieceFromIds = new Set(pieces.pairs.map((pair) => pair.fromElement.id));
	const pieceToIds = new Set(pieces.pairs.map((pair) => pair.toElement.id));
	const wholePairs = match.pairs.filter((pair) => !piecePairs.has(pair));
	const ghostElements = flattenedOutgoing
		.filter((element) => ghostIds.has(element.id) && !pieceFromIds.has(element.id))
		.map((element) => {
			const target = offstageRestyles.get(element.id);
			if (!target || element.type !== 'text' || target.type !== 'text') return element;
			return {
				...element,
				width: target.width,
				height: target.height,
				textStyle: target.textStyle,
				textSegments: target.textSegments,
			} as PptxElement;
		});

	const lifted = resolveMorphLiftedIds({
		mode,
		measureText: options.measureText,
		pairs: match.pairs,
		wholePairs,
		ghostIds,
		flattenedOutgoing,
		flattenedIncoming,
		pieceToIds,
		risingIds: bodies.rising,
		animated: (id) => incomingAnimations.has(id),
	});
	const overlayIncomingAnimations = settleMorphLifts({
		durationMs,
		lifted,
		pieces,
		pieceFromIds,
		pieceToIds,
		incomingAnimations,
		outgoingAnimations,
		outlineAnimations,
		outlineStrandAnimations,
		customGeometryAnimations,
		keyframes,
	});
	const overlay = assembleMorphOverlay({
		animations,
		pairs: match.pairs,
		wholePairs,
		ghostIds,
		ghostElements,
		flattenedOutgoing,
		flattenedIncoming,
		pieces,
		outgoingAnimations,
		overlayIncomingAnimations,
	});

	return {
		keyframesCss: keyframes.join('\n'),
		incomingAnimations,
		outgoingAnimations,
		incomingImageAnimations: channels.incomingImageAnimations,
		outgoingImageAnimations: channels.outgoingImageAnimations,
		overlayIncomingAnimations,
		outgoingElements: overlay.outgoingElements,
		outgoingTextReplacements: new Map(
			[...offstageRestyles]
				.filter(([id]) => ghostIds.has(id))
				.map(([id, target]) => [id, target.id]),
		),
		overlayIncomingElements: overlay.overlayIncomingElements,
		overlayLayers: overlay.overlayLayers,
		overlayStack: overlay.overlayStack,
		crossfadeGroups: overlay.crossfadeGroups,
		...(pieces.sources.size > 0 ? { textPieceSources: pieces.sources } : {}),
		...(bodies.parts.size > 0 ? { morphCopyParts: bodies.parts } : {}),
		...(outlineAnimations.size > 0 ? { outlineAnimations } : {}),
		...(outlineStrandAnimations.size > 0 ? { outlineStrandAnimations } : {}),
		...(customGeometryAnimations.size > 0 ? { customGeometryAnimations } : {}),
		...(channels.lineAnimations.size > 0 ? { lineAnimations: channels.lineAnimations } : {}),
		durationMs,
	};
}
