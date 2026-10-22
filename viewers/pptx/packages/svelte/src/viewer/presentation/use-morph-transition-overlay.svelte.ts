import {
	buildMorphScopedCss,
	buildMorphTransitionPlan,
	MORPH_CROSSFADE_GROUP_CSS_TEXT,
	MORPH_CROSSFADE_HALF_BLEND_MODE,
	morphOptionToMode,
} from 'ooxml-ui/pptx';
import type { PptxSlide, PptxSlideTransition } from 'pptx-viewer-core';

/**
 * Morph-specific reactive state for `PresentationTransitionOverlay.svelte`,
 * split out to keep that file under the project's per-file LOC budget.
 *
 * Morph is not a whole-slide wipe: individual shapes travel from where they
 * sat on the outgoing slide to where they sit on the incoming one, so when
 * the transition is `morph` the overlay's two stacked layers are re-purposed
 * - the incoming layer plays per-element keyframes (scoped by
 * `data-pptx-morph-incoming`), and the outgoing layer paints a moving copy of
 * the outgoing slide, each shape gliding onto its counterpart (dissolving
 * into it when its appearance changed) or fading out in place when it has
 * none.
 *
 * Inputs are getters (not raw values) so the `$derived`s stay reactive to the
 * host component's own props; call this once, at component init, per the
 * Svelte 5 rules-of-runes.
 */
export function useMorphTransitionOverlay(options: {
	transition: () => PptxSlideTransition | undefined;
	outgoingSlide: () => PptxSlide | undefined;
	incomingSlide: () => PptxSlide | undefined;
	durationMs: () => number;
}) {
	const morphPlan = $derived(
		options.transition()?.type === 'morph'
			? buildMorphTransitionPlan(
					options.outgoingSlide(),
					options.incomingSlide(),
					options.durationMs(),
					morphOptionToMode((options.transition() as PptxSlideTransition).morphOption),
				)
			: undefined,
	);

	/**
	 * The ghosts and lifted arrivals, one stacked layer per run of the plan's
	 * `overlayLayers`, back to front above the incoming layer (z-index 1).
	 *
	 * A ghost run renders against the outgoing slide inside a
	 * `data-pptx-morph-outgoing` scope; a lifted run (arriving shapes a ghost
	 * would otherwise hide, issue #146) against the incoming slide inside a
	 * `data-pptx-morph-lifted` one, its copy on the incoming layer held
	 * invisible by the plan. One shared order, rather than every ghost under
	 * every lifted shape, is what keeps a full-slide picture fading in from
	 * veiling the title travelling above it (issue #141).
	 */
	const morphOverlayLayers = $derived.by(() => {
		const outgoing = options.outgoingSlide();
		const incoming = options.incomingSlide();
		if (!morphPlan || !outgoing) {
			return [];
		}
		return morphPlan.overlayLayers.flatMap((layer, index) => {
			const slide = layer.side === 'outgoing' ? outgoing : incoming;
			if (!slide) {
				return [];
			}
			return [
				{
					key: `${layer.side}-${layer.elements[0]?.id ?? index}`,
					side: layer.side,
					style: `z-index: ${MORPH_LAYER_Z_BASE + index};`,
					slide: { ...slide, elements: [...layer.elements] } as PptxSlide,
				},
			];
		});
	});

	/** One half of a grouped pair: blends additively, and carries the dissolve. */
	function halfStyle(animation: string | undefined): string {
		return `mix-blend-mode: ${MORPH_CROSSFADE_HALF_BLEND_MODE};${
			animation === undefined ? '' : ` animation: ${animation};`
		}`;
	}

	/**
	 * The pairs the overlay paints BOTH halves of, each as one isolated group so
	 * the halves are summed rather than stacked: two source-over fades leave the
	 * ink they share at 0.75 of full strength mid-transition, biting chunks out
	 * of glyphs that cross during a text dissolve, where PowerPoint's own blend
	 * keeps the two coefficients summing to 1.0 (issue #161).
	 */
	const morphCrossfadeGroups = $derived.by(() => {
		const outgoingSlide = options.outgoingSlide();
		const incomingSlide = options.incomingSlide();
		if (!morphPlan || !outgoingSlide || !incomingSlide) {
			return [];
		}
		return morphPlan.crossfadeGroups.map((group, index) => ({
			key: group.incoming.id,
			// `isolation` makes the group a stacking context, so it carries its
			// own z-index to stay above every overlay layer.
			style: `${MORPH_CROSSFADE_GROUP_CSS_TEXT} z-index: ${
				MORPH_LAYER_Z_BASE + morphPlan.overlayLayers.length + index
			};`,
			// The dissolve rides these WRAPPERS, not the elements: a pair
			// dissolving in place never moves, and an animation on the small
			// element box gives it a compositing layer whose raster snaps to
			// whole device pixels, painting the wording a fraction of a pixel
			// off the live stage (issue #161).
			outgoingStyle: halfStyle(group.outgoingAnimation),
			incomingStyle: halfStyle(group.incomingAnimation),
			outgoing: { ...outgoingSlide, elements: [group.outgoing] },
			incoming: { ...incomingSlide, elements: [group.incoming] },
		}));
	});

	const morphCss = $derived(
		morphPlan
			? [
					buildMorphScopedCss(morphPlan, 'data-pptx-morph-incoming', 'incoming'),
					buildMorphScopedCss(morphPlan, 'data-pptx-morph-outgoing', 'outgoing'),
					buildMorphScopedCss(morphPlan, 'data-pptx-morph-lifted', 'lifted'),
				].join('\n')
			: '',
	);

	return {
		get morphPlan() {
			return morphPlan;
		},
		get morphOverlayLayers() {
			return morphOverlayLayers;
		},
		get morphCrossfadeGroups() {
			return morphCrossfadeGroups;
		},
		get morphCss() {
			return morphCss;
		},
	};
}

/** The z-index of the first overlay layer, just above the incoming layer (1). */
const MORPH_LAYER_Z_BASE = 2;

/** One entry of {@link useMorphTransitionOverlay}'s `morphOverlayLayers`. */
export type MorphOverlayLayerView = ReturnType<
	typeof useMorphTransitionOverlay
>['morphOverlayLayers'][number];

/** One entry of {@link useMorphTransitionOverlay}'s `morphCrossfadeGroups`. */
export type MorphCrossfadeGroupView = ReturnType<
	typeof useMorphTransitionOverlay
>['morphCrossfadeGroups'][number];
