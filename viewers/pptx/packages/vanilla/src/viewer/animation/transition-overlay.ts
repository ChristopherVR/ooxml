import {
	buildMorphScopedCss,
	buildMorphTransitionPlan,
	getFragmentedTransitionDescriptor,
	morphOptionToMode,
	resolveSlideTransition,
	resolveTransitionDurationMs,
} from 'ooxml-ui/pptx';
import type { PptxSlide, PptxSlideTransition } from 'pptx-viewer-core';

import { ensurePresentationKeyframes } from './animation-dom';
import { buildFragmentedLayer } from './fragmented-transition-layer';
import { buildMorphCrossfadeGroups, buildMorphOverlayLayers } from './morph-transition-layers';
import { buildLayer } from './transition-layer';

export interface TransitionOverlayParams {
	doc: Document;
	/** The stage host the overlay is layered into (position: relative). */
	stageWrap: HTMLElement;
	/** The detached outgoing (previous) stage node, re-attached as a snapshot. */
	outgoing: HTMLElement;
	/** A fully-visible clone of the incoming (new) stage node. */
	incoming: HTMLElement;
	/** The incoming slide's transition definition. */
	transition: PptxSlideTransition;
	/** The outgoing slide model. Required to play a Morph transition. */
	outgoingSlide?: PptxSlide;
	/** The incoming slide model. Required to play a Morph transition. */
	incomingSlide?: PptxSlide;
	/** Called once the transition duration elapses (or the overlay is cancelled). */
	onDone: () => void;
}

/**
 * Play a slide-change transition in presentation mode by stacking the outgoing
 * snapshot and the incoming clone as two absolutely-positioned layers over the
 * stage, each driven by the shared `resolveSlideTransition` CSS `animation`
 * shorthand. After the resolved duration the overlay removes itself and calls
 * `onDone`.
 *
 * Returns a cancel function that tears the overlay down immediately (used when
 * the stage is rebuilt again before the transition finishes).
 */
export function playTransitionOverlay(params: TransitionOverlayParams): () => void {
	const { doc, stageWrap, outgoing, incoming, transition } = params;
	ensurePresentationKeyframes(doc);

	const resolved = resolveSlideTransition(transition);
	const durationMs = resolveTransitionDurationMs(transition);

	// Multi-fragment descriptor for the seven cinematic transitions measured
	// as many independent fragments/particles/panels (vortex, honeycomb,
	// glitter, shred, fracture, curtains, airplane). `undefined` for every
	// other type, in which case `resolved` above (the single-layer resolver)
	// drives both layers exactly as before.
	const fragmented = getFragmentedTransitionDescriptor(
		transition.type,
		durationMs,
		transition.direction,
		transition.spokes,
		transition.pattern,
	);

	const overlay = doc.createElement('div');
	overlay.className = 'pptxv-transition-overlay';
	// Neutral marker every other binding already emits, so a product e2e can
	// find the overlay without naming this binding's class (the neutrality
	// check in `scripts/check-e2e-neutrality.mjs` rejects a `.pptxv-` selector).
	overlay.dataset.pptxTransitionOverlay = '';
	overlay.dataset.pptxTransitionOverlay = '';
	overlay.style.position = 'absolute';
	overlay.style.inset = '0';
	overlay.style.overflow = 'hidden';
	overlay.style.pointerEvents = 'none';
	overlay.style.zIndex = '30';

	// Morph moves individual shapes between the two slides instead of wiping the
	// whole surface. When a plan is available the layers stay unanimated (a
	// layer-wide animation would drag every shape as one block), the incoming
	// layer plays per-element keyframes, and every shape in the outgoing snapshot
	// glides onto its counterpart - dissolving into it when its appearance
	// changed, or fading out in place when it has none.
	const morphPlan =
		transition.type === 'morph'
			? buildMorphTransitionPlan(
					params.outgoingSlide,
					params.incomingSlide,
					durationMs,
					morphOptionToMode(transition.morphOption),
				)
			: undefined;

	const outgoingZIndex = resolved.outgoingOnTop ? 2 : 1;
	const incomingZIndex = morphPlan ? 1 : resolved.outgoingOnTop ? 1 : 2;
	const morphLayers = buildMorphOverlayLayers(doc, outgoing, incoming, morphPlan);
	const outLayer =
		!morphPlan && fragmented?.outgoing
			? buildFragmentedLayer(doc, outgoing, fragmented.outgoing, outgoingZIndex, 'outgoing')
			: buildLayer(doc, outgoing, outgoingZIndex, resolved.outgoing, 'outgoing');
	const inLayer =
		!morphPlan && fragmented?.incoming
			? buildFragmentedLayer(doc, incoming, fragmented.incoming, incomingZIndex, 'incoming')
			: buildLayer(
					doc,
					incoming,
					incomingZIndex,
					morphPlan ? 'none' : resolved.incoming,
					'incoming',
				);

	// The ghosts, and the shapes ARRIVING above one that would otherwise be
	// drawn under it on the incoming layer and never seen dissolving in (issue
	// #146), in the plan's single back-to-front order: one layer per run, each
	// a background-free copy of its slide (issue #141). A pair the overlay
	// paints BOTH halves of goes in its own isolated group above them. See
	// `morph-transition-layers.ts`; both are built from clones taken before
	// the stages are wrapped into their layers.
	const crossfadeGroups = buildMorphCrossfadeGroups(doc, outgoing, incoming, morphPlan);

	if (morphPlan) {
		inLayer.dataset.pptxMorphIncoming = 'true';
		const style = doc.createElement('style');
		style.textContent = [
			buildMorphScopedCss(morphPlan, 'data-pptx-morph-incoming', 'incoming'),
			buildMorphScopedCss(morphPlan, 'data-pptx-morph-outgoing', 'outgoing'),
			buildMorphScopedCss(morphPlan, 'data-pptx-morph-lifted', 'lifted'),
		].join('\n');
		overlay.appendChild(style);
	}

	// A morph paints its ghosts in `morphLayers`, never the whole snapshot.
	if (morphPlan) {
		overlay.appendChild(inLayer);
	} else {
		overlay.append(outLayer, inLayer);
	}
	for (const layer of morphLayers) {
		overlay.appendChild(layer);
	}
	for (const group of crossfadeGroups) {
		overlay.appendChild(group);
	}
	stageWrap.appendChild(overlay);

	let done = false;
	const finish = (): void => {
		if (done) {
			return;
		}
		done = true;
		clearTimeout(timer);
		overlay.remove();
		params.onDone();
	};

	// A small buffer past the animation duration lets the CSS `forwards` fill
	// settle before the host reveals the static incoming stage underneath.
	const timer = setTimeout(finish, Math.max(0, durationMs) + 50);

	return finish;
}
