/**
 * Assemble the overlay of a plan: hand the lifted arrivals and the piece copies
 * their animations, group the cross-dissolving pairs, and stack everything the
 * overlay paints back to front (see {@link MorphTransitionPlan.overlayStack}).
 *
 * @module render/morph-plan-overlay
 */
import type { PptxElement } from 'ooxml-core/pptx';

import type { MorphCrossfadeGroup } from './morph-crossfade-group';
import { resolveMorphCrossfadeGroups } from './morph-crossfade-group';
import type { MorphOverlayLayer } from './morph-overlay-layers';
import type { MorphOverlayStackEntry } from './morph-overlay-stack';
import { resolveMorphOverlayStack } from './morph-overlay-stack';
import { liftCrossfadeToWrapper } from './morph-plan-crossfade';
import type { MorphGeneratedAnimations } from './morph-plan-channels';
import type { MorphTextPiecePlan } from './morph-text-piece-plan';
import type { MorphPair } from './morph-types';

/**
 * Keyframes for an incoming shape whose dissolve has been lifted into the
 * overlay: the copy left on the live stage holds at nothing for the whole
 * morph, so the two copies never composite with each other.
 */
const LIFTED_HIDDEN_NAME = 'pptx-morph-lifted-hidden';
const LIFTED_HIDDEN_KEYFRAMES = `
@keyframes ${LIFTED_HIDDEN_NAME} {
\tfrom {
\t\topacity: 0;
\t}
\tto {
\t\topacity: 0;
\t}
}`;

/** The maps {@link settleMorphLifts} reads and edits in place. */
export interface MorphSettleInput {
	readonly durationMs: number;
	readonly lifted: ReadonlySet<string>;
	readonly pieces: MorphTextPiecePlan;
	readonly pieceFromIds: ReadonlySet<string>;
	readonly pieceToIds: ReadonlySet<string>;
	readonly incomingAnimations: Map<string, string>;
	readonly outgoingAnimations: Map<string, string>;
	readonly outlineAnimations: Map<string, string>;
	readonly outlineStrandAnimations: Map<string, string[]>;
	readonly customGeometryAnimations: Map<string, string>;
	/** The plan's keyframes; the hold and the pieces' blocks are appended. */
	readonly keyframes: string[];
}

/**
 * Move each lifted arrival's animation to the overlay and hold its stage copy
 * invisible; hold the piece pairs' stage copies too and drop the animations of
 * the ghosts and outlines the pieces replace.
 *
 * @returns The animations of the copies the overlay paints above the ghosts.
 */
export function settleMorphLifts(input: MorphSettleInput): Map<string, string> {
	const { durationMs, incomingAnimations, outgoingAnimations, pieceFromIds, pieceToIds } = input;
	const hold = `${LIFTED_HIDDEN_NAME} ${durationMs}ms linear forwards`;
	const overlayIncomingAnimations = new Map<string, string>();
	for (const id of input.lifted) {
		const animation = incomingAnimations.get(id);
		if (animation === undefined) {
			continue;
		}
		overlayIncomingAnimations.set(id, animation);
		incomingAnimations.set(id, hold);
	}
	for (const id of pieceToIds) {
		incomingAnimations.set(id, hold);
	}
	for (const id of pieceFromIds) {
		outgoingAnimations.delete(id);
		input.outlineAnimations.delete(id);
		input.outlineStrandAnimations.delete(id);
		input.customGeometryAnimations.delete(id);
	}
	for (const id of pieceToIds) {
		input.outlineAnimations.delete(id);
		input.outlineStrandAnimations.delete(id);
		input.customGeometryAnimations.delete(id);
	}
	if (overlayIncomingAnimations.size > 0 || pieceToIds.size > 0) {
		input.keyframes.push(LIFTED_HIDDEN_KEYFRAMES);
	}
	input.keyframes.push(...input.pieces.keyframes);
	return overlayIncomingAnimations;
}

/** Everything {@link assembleMorphOverlay} reads. */
export interface MorphOverlayInput {
	readonly animations: MorphGeneratedAnimations;
	readonly pairs: readonly MorphPair[];
	readonly wholePairs: readonly MorphPair[];
	readonly ghostIds: ReadonlySet<string>;
	readonly ghostElements: readonly PptxElement[];
	readonly flattenedOutgoing: readonly PptxElement[];
	readonly flattenedIncoming: readonly PptxElement[];
	readonly pieces: MorphTextPiecePlan;
	readonly outgoingAnimations: Map<string, string>;
	/** Edited in place: the pieces' incoming animations join it. */
	readonly overlayIncomingAnimations: Map<string, string>;
}

/** What the overlay paints and in which order. */
export interface MorphOverlay {
	crossfadeGroups: MorphCrossfadeGroup[];
	outgoingElements: PptxElement[];
	overlayIncomingElements: PptxElement[];
	overlayLayers: MorphOverlayLayer[];
	overlayStack: MorphOverlayStackEntry[];
}

/** Group, filter and stack the overlay's contents. */
export function assembleMorphOverlay(input: MorphOverlayInput): MorphOverlay {
	const { pieces, flattenedOutgoing, flattenedIncoming, overlayIncomingAnimations } = input;
	// A lifted half whose ghost is painted too is one end of a cross-dissolve
	// with both ends in this overlay, so the two are handed over as a pair and
	// taken out of the flat layers. Stacking them there composites them
	// source-over, which dips their shared ink toward the backdrop instead of
	// summing it the way PowerPoint's own blend does (issue #161).
	const crossfadeGroups = resolveMorphCrossfadeGroups(
		input.wholePairs,
		input.ghostIds,
		new Set(overlayIncomingAnimations.keys()),
		flattenedIncoming,
	).map((group) =>
		liftCrossfadeToWrapper(
			group,
			input.animations,
			input.outgoingAnimations,
			overlayIncomingAnimations,
		),
	);
	crossfadeGroups.push(...pieces.groups);
	for (const [id, animation] of pieces.outgoingAnimations) {
		input.outgoingAnimations.set(id, animation);
	}
	for (const [id, animation] of pieces.incomingAnimations) {
		overlayIncomingAnimations.set(id, animation);
	}
	const groupedOutgoingIds = new Set(crossfadeGroups.map((group) => group.outgoing.id));
	const groupedIncomingIds = new Set(crossfadeGroups.map((group) => group.incoming.id));
	const outgoingElements = input.ghostElements.filter(
		(element) => !groupedOutgoingIds.has(element.id),
	);
	const overlayIncomingElements = flattenedIncoming.filter(
		(element) => overlayIncomingAnimations.has(element.id) && !groupedIncomingIds.has(element.id),
	);
	const { layers: overlayLayers, stack: overlayStack } = resolveMorphOverlayStack(
		outgoingElements,
		overlayIncomingElements,
		crossfadeGroups,
		flattenedOutgoing,
		flattenedIncoming,
		input.pairs,
		pieces.sources,
	);
	return {
		crossfadeGroups,
		outgoingElements,
		overlayIncomingElements,
		overlayLayers,
		overlayStack,
	};
}
