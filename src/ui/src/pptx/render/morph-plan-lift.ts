/**
 * Which incoming shapes the overlay paints above the ghosts instead of leaving
 * them on the live stage (see {@link MorphTransitionPlan.overlayIncomingElements}).
 *
 * @module render/morph-plan-lift
 */
import type { PptxElement } from 'ooxml-core/pptx';

import { morphPairIncomingFadesIn, morphPairNeedsCrossfade } from './morph-animation';
import { resolveMorphOverlayArrivals } from './morph-overlay-order';
import { resolveMorphSoftEdgeLifts } from './morph-soft-edge-blend';
import { resolveMorphTextBlendLifts } from './morph-text-blend';
import type { MorphTextMeasure } from './morph-text-frame';
import type { MorphMode, MorphPair } from './morph-types';

/** Everything {@link resolveMorphLiftedIds} reads about the transition. */
export interface MorphLiftInput {
	readonly mode: MorphMode;
	readonly measureText: MorphTextMeasure | undefined;
	/** Every matched pair. */
	readonly pairs: readonly MorphPair[];
	/** The pairs not carried piece by piece. */
	readonly wholePairs: readonly MorphPair[];
	readonly ghostIds: ReadonlySet<string>;
	readonly flattenedOutgoing: readonly PptxElement[];
	readonly flattenedIncoming: readonly PptxElement[];
	/** Incoming ids of the pairs carried piece by piece. */
	readonly pieceToIds: ReadonlySet<string>;
	/** Incoming ids a body carry raises above the ghosts. */
	readonly risingIds: Iterable<string>;
	/** Whether the plan holds a container animation for an incoming id. */
	readonly animated: (id: string) => boolean;
}

/**
 * The incoming ids the overlay paints above the ghosts.
 *
 * Everything the overlay paints hides whatever the live stage is doing
 * underneath, which is wrong for a shape that ARRIVES on top of a ghost:
 * it dissolves in where nobody can see it and appears in one frame when the
 * overlay is torn down (issue #146 - the wheel's centre disc is unchanged,
 * so its opaque ghost sat over the new title, body and button for the whole
 * morph). Those few move up into the overlay, above the ghosts, and the
 * copy on the stage is held invisible so the two never composite.
 *
 * Only a ghost that KEEPS its opacity counts. One that dissolves is out of
 * the way inside the first quarter, long before an arrival begins to appear,
 * so it hides nothing worth moving an animation for.
 */
export function resolveMorphLiftedIds(input: MorphLiftInput): Set<string> {
	const { pairs, wholePairs, ghostIds, flattenedOutgoing, flattenedIncoming, pieceToIds } = input;
	const holdingGhostIds = new Set(
		wholePairs
			.filter(
				(candidate) =>
					ghostIds.has(candidate.fromElement.id) &&
					!morphPairNeedsCrossfade(candidate.fromElement, candidate.toElement),
			)
			.map((candidate) => candidate.fromElement.id),
	);
	// A matched pair's incoming half can be hidden by a holding ghost just as
	// easily as an arrival can - the wheel deck dissolves its centre wording
	// inside an unchanged opaque disc (issue #160) - but only one that DISSOLVES
	// IN may be lifted over its own ghost. One pinned at full strength has to
	// stay underneath it, or the crossfade becomes a cut.
	const dissolvingInIds = new Set(
		wholePairs
			.filter((candidate) =>
				morphPairIncomingFadesIn(
					candidate.fromElement,
					candidate.toElement,
					ghostIds.has(candidate.fromElement.id),
				),
			)
			.map((candidate) => candidate.toElement.id),
	);
	const lifted = resolveMorphOverlayArrivals(
		flattenedOutgoing,
		flattenedIncoming,
		pairs,
		holdingGhostIds,
		dissolvingInIds,
	);
	// Restyled text whose two halves are carried glyph over glyph dissolves
	// additively, as PowerPoint's does, with whatever the incoming slide stacks
	// above it raised along with it (see `morph-text-blend`).
	if (input.mode === 'object') {
		for (const id of resolveMorphTextBlendLifts(
			{
				outgoing: flattenedOutgoing,
				incoming: flattenedIncoming,
				pairs,
				ghostIds,
				overlaid: new Set([...lifted, ...pieceToIds]),
				animated: input.animated,
			},
			dissolvingInIds,
			input.measureText,
		)) {
			lifted.add(id);
		}
	}
	// A soft-edged pair dissolves its two feathers additively as well.
	for (const id of resolveMorphSoftEdgeLifts(
		{
			outgoing: flattenedOutgoing,
			incoming: flattenedIncoming,
			pairs: wholePairs,
			ghostIds,
			overlaid: new Set([...lifted, ...pieceToIds]),
			animated: input.animated,
		},
		dissolvingInIds,
	)) {
		lifted.add(id);
	}
	for (const id of input.risingIds) {
		lifted.add(id);
	}
	for (const id of pieceToIds) {
		lifted.delete(id);
	}
	return lifted;
}
