/**
 * The contract of a morph plan: {@link MorphTransitionPlan} (what a binding
 * plays) and {@link MorphPlanOptions} (the per-binding knobs of building one).
 * `morph-plan` re-exports both and holds the entry point; see its header for
 * how a morph is composed.
 *
 * @module render/morph-plan-types
 */
import type { PptxElement } from 'ooxml-core/pptx';

import type { MorphCrossfadeGroup } from './morph-crossfade-group';
import type { MorphViewSize } from './morph-flip-turn';
import type { MorphOverlayLayer } from './morph-overlay-layers';
import type { MorphOverlayStackEntry } from './morph-overlay-stack';
import type { MorphCopyPart } from './morph-text-body';
import type { MorphTextMeasure } from './morph-text-frame';

/** Everything needed to play one morph transition. */
export interface MorphTransitionPlan {
	/** All `@keyframes` blocks, to inject as a single stylesheet. */
	keyframesCss: string;
	/** Incoming-slide element id -> CSS `animation` shorthand. */
	incomingAnimations: Map<string, string>;
	/** Outgoing-slide element id -> CSS `animation` shorthand. */
	outgoingAnimations: Map<string, string>;
	/**
	 * Incoming-slide element id -> CSS `animation` shorthand for the `<img>`
	 * INSIDE that element, never for the element's own container.
	 *
	 * A picture's source crop (`a:srcRect`, which is what PowerPoint's "Scale
	 * Height"/"Scale Width" writes) is painted by transforming the img within
	 * its frame, and the frame itself is usually identical on both slides - so
	 * this is a separate channel rather than another entry in
	 * {@link MorphTransitionPlan.incomingAnimations}, which would collide on the
	 * same element id. {@link buildMorphScopedCss} emits it as a descendant
	 * rule; a binding that applies animations as inline props instead must
	 * target the img itself (issue #148).
	 */
	incomingImageAnimations: Map<string, string>;
	/** Outgoing (ghost) counterpart of {@link MorphTransitionPlan.incomingImageAnimations}. */
	outgoingImageAnimations: Map<string, string>;
	/**
	 * Element id -> CSS `animation` shorthand for the PATH of the element's
	 * outline overlay (the `<svg data-pptx-outline>` every binding paints a
	 * centred or gradient outline with), present only when a shape's outline is
	 * re-pathed per frame (see `morph-shape-carry`). Like the image channel it is
	 * a descendant rule, emitted by {@link buildMorphScopedCss}; the ids may be
	 * incoming, outgoing or copy ids, so each half's rules carry all of them.
	 */
	outlineAnimations?: Map<string, string>;
	/**
	 * Element id -> the `animation` shorthand of each strand path of a compound
	 * outline, in paint order (each strand is re-pathed on its own); emitted
	 * next to {@link MorphTransitionPlan.outlineAnimations}.
	 */
	outlineStrandAnimations?: Map<string, string[]>;
	/**
	 * Element id -> CSS `animation` shorthand for the paths of a freeform's
	 * custom-geometry `<svg data-pptx-custom-geometry>`, which a binding that
	 * paints custom geometry as a vector (rather than a clipped box) stamps;
	 * emitted next to {@link MorphTransitionPlan.outlineAnimations}.
	 */
	customGeometryAnimations?: Map<string, string>;
	/**
	 * Element id -> CSS `animation` shorthand for the painted strand
	 * (`[data-pptx-line]`) of a straight connector, which keeps its width while
	 * the wrapper's journey scales (see `morph-line-carry`); emitted next to
	 * {@link MorphTransitionPlan.outlineAnimations}.
	 */
	lineAnimations?: Map<string, string>;
	/**
	 * Incoming-slide element id -> CSS `animation` shorthand for the copy the
	 * overlay paints, above every ghost.
	 *
	 * These are the arriving shapes a ghost would otherwise hide for the whole
	 * transition (see {@link MorphTransitionPlan.overlayIncomingElements}). Their
	 * entry in {@link MorphTransitionPlan.incomingAnimations} has been replaced
	 * by one that holds them invisible, so the stage copy stays out of the way
	 * and only this one is seen.
	 */
	overlayIncomingAnimations: Map<string, string>;
	/**
	 * The incoming-slide elements to paint in the overlay ON TOP of
	 * {@link MorphTransitionPlan.outgoingElements}, in document order.
	 *
	 * The overlay is one flat layer above the live stage, which is only faithful
	 * while every ghost really does belong on top of everything the stage draws.
	 * A shape arriving INSIDE a persisting one does not: the wheel deck's centre
	 * disc is unchanged between slides, so its ghost is opaque for the whole
	 * morph, and the title and body dissolving in within it were invisible until
	 * the overlay came down (issue #146). Painting those here restores the order
	 * PowerPoint composites in.
	 *
	 * Usually empty. A binding renders these with the INCOMING slide as their
	 * context, applying {@link MorphTransitionPlan.overlayIncomingAnimations}.
	 *
	 * A lifted half that is one end of a cross-dissolve is NOT here: it moves
	 * into {@link MorphTransitionPlan.crossfadeGroups} so it can be composited
	 * with its own ghost.
	 */
	overlayIncomingElements: PptxElement[];
	/**
	 * Pairs the overlay paints BOTH halves of, to be rendered as one isolated
	 * group blending additively (see `morph-crossfade-group`).
	 *
	 * Their elements are deliberately absent from
	 * {@link MorphTransitionPlan.outgoingElements} and
	 * {@link MorphTransitionPlan.overlayIncomingElements}: a binding that renders
	 * the groups paints each half exactly once, and one that ignored them would
	 * drop the pair rather than paint it twice.
	 *
	 * A pair that dissolves where it stands also has its dissolve moved OFF the
	 * two elements and onto {@link MorphCrossfadeGroup.outgoingAnimation} /
	 * {@link MorphCrossfadeGroup.incomingAnimation}, for the binding to put on
	 * the wrapper it paints each half in; the elements are then absent from
	 * `outgoingAnimations` / `overlayIncomingAnimations` too, so no scoped rule
	 * animates them. One that MOVES keeps its animations where they were.
	 *
	 * Painted above {@link MorphTransitionPlan.overlayIncomingElements}, in the
	 * incoming slide's document order.
	 */
	crossfadeGroups: MorphCrossfadeGroup[];
	/**
	 * The outgoing slide's elements, in document order, for the binding to
	 * render in its transition overlay for the duration of the morph. Most carry
	 * an entry in {@link MorphTransitionPlan.outgoingAnimations}: they either
	 * fade out in place (no counterpart) or glide onto a counterpart, dissolving
	 * into it when the appearance changed.
	 *
	 * A ghost standing in for an INERT pair deliberately has no entry there and
	 * is painted statically. Its keyframes would run from itself to itself, and
	 * a running animation would put it on its own compositing layer, whose
	 * raster the browser snaps to whole device pixels - visibly shifting a shape
	 * that is not supposed to move at all (issue #161).
	 *
	 * This is a SUBSET of the outgoing slide: a shape the live stage already
	 * draws identically is left out, because the overlay is opaque above it and
	 * would hide the incoming slide's own arrivals. See `resolveMorphGhostIds`.
	 */
	outgoingElements: PptxElement[];
	/** Outgoing ghost id to the incoming text layout it paints while travelling. */
	outgoingTextReplacements?: ReadonlyMap<string, string>;
	/**
	 * {@link MorphTransitionPlan.outgoingElements} and
	 * {@link MorphTransitionPlan.overlayIncomingElements} together, in the ONE
	 * back-to-front order the overlay must paint them in, as runs of
	 * consecutive same-slide shapes (see `morph-overlay-layers`).
	 *
	 * A binding renders each run as its own stacked layer, in this order: an
	 * `'outgoing'` run against the outgoing slide inside a
	 * `data-pptx-morph-outgoing` scope, an `'incoming'` run against the incoming
	 * slide inside a `data-pptx-morph-lifted` scope. Painting every ghost below
	 * every lifted arrival instead let a full-slide picture fading in veil the
	 * title travelling above it (issue #141).
	 *
	 * A run is split wherever a crossfade group belongs inside it, and
	 * {@link MorphTransitionPlan.overlayStack} says where each group goes.
	 */
	overlayLayers: MorphOverlayLayer[];
	/**
	 * Every run of {@link MorphTransitionPlan.overlayLayers} and every
	 * {@link MorphTransitionPlan.crossfadeGroups} group, back to front. Read it
	 * through `morphOverlayStackSlots` or `morphOverlayElementSlots` (see
	 * `morph-overlay-stack`). Painting every group above every run let three
	 * dissolving text columns show through the picture bands the incoming slide
	 * stacks above them, then vanish from them when the overlay came down
	 * (issue #141, slides 6 to 7).
	 */
	overlayStack: MorphOverlayStackEntry[];
	/**
	 * Id -> source id for the copies the overlay paints of a text pair that is
	 * carried piece by piece (see `morph-text-piece-plan`). Each copy is in
	 * {@link MorphTransitionPlan.crossfadeGroups} as an element of its own; a
	 * binding that builds its overlay from a clone of the rendered slide instead
	 * clones the source's node under the copy's id.
	 */
	textPieceSources?: ReadonlyMap<string, string>;
	/**
	 * Copy id -> the part of its source a copy paints, for a shape whose text is
	 * carried apart from its body (see `morph-text-body`): `'body'` without the
	 * text, `'text'` without the fill and outline. The copy's element already
	 * says so; a binding that clones the source's node strips the other part.
	 */
	morphCopyParts?: ReadonlyMap<string, MorphCopyPart>;
	/** Animation duration in ms, echoed for convenience. */
	durationMs: number;
}

/** Per-binding knobs for {@link buildMorphTransitionPlan}. */
export interface MorphPlanOptions {
	/**
	 * The z-index the live stage gives the incoming slide's FIRST top-level
	 * element. Stacking-order swaps ride z-index journeys written in the
	 * stage's real z space, so a binding that paints master/layout shapes
	 * beneath the slide (React's presentation stage) passes their count;
	 * a stage that starts the slide at 0 leaves it out.
	 */
	readonly zIndexBase?: number;
	/**
	 * How text is measured when a pair's unchanged text is carried line for
	 * line (see `morph-text-frame`). Defaults to a canvas; tests pass their own.
	 */
	readonly measureText?: MorphTextMeasure;
	/**
	 * The slide's size in element pixels: the camera a turning flip is seen
	 * through sits over its centre (see `morph-flip-turn`). Defaults to the
	 * 16:9 slide, 1280 x 720.
	 */
	readonly slideSize?: MorphViewSize;
}
