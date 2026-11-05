/**
 * The overlay copies that carry a rewrapped text pair piece by piece (see
 * `morph-text-pieces` for what PowerPoint does).
 *
 * The overlay can only move a whole element, so each piece is painted as its
 * own copy of BOTH halves, clipped to the band of the text that piece owns: a
 * copy of the outgoing box showing only those words on the old layout, and a
 * copy of the incoming box showing only them on the new one. The two copies of
 * a piece are one crossfade group, so they sum the way a matched pair's
 * dissolve does, and travel together on the piece's own mapping. The pair's
 * own ghost is not painted and its stage copy is held invisible: the pieces
 * are the whole of it.
 *
 * A copy is the half's element under a new id, so every binding that renders
 * a group from its element paints it unchanged. One that clones its own DOM
 * instead finds the element to clone through
 * {@link MorphTextPiecePlan.sources}.
 *
 * @module render/morph-text-piece-plan
 */
import type { PptxElement } from 'ooxml-core/pptx';

import type { MorphCrossfadeGroup } from './morph-crossfade-group';
import type { MorphTextMeasure } from './morph-text-frame';
import {
	morphTextGhostEnd,
	morphTextIncomingStart,
	resolveMorphTextAlignment,
} from './morph-text-frame';
import type { MorphTextBand } from './morph-text-bands';
import { morphTextBandOutline } from './morph-text-bands';
import { resolveMorphTextPieces } from './morph-text-pieces';
import {
	morphOffstageTextAdoptsTarget,
	morphOffstageTextSizeChange,
	morphTextReplacedInSlot,
} from './morph-text-slot';
import type { MorphPair } from './morph-types';
import { MORPH_EASING } from './morph-types';

/** What the plan adds for the pairs carried piece by piece. */
export interface MorphTextPiecePlan {
	/** The pairs replaced by pieces. */
	pairs: MorphPair[];
	/** One group per piece, in the pairs' order. */
	groups: MorphCrossfadeGroup[];
	/** Copy id -> CSS `animation` shorthand, for the outgoing and incoming copies. */
	outgoingAnimations: Map<string, string>;
	incomingAnimations: Map<string, string>;
	keyframes: string[];
	/**
	 * Copy id -> the id of the element it copies. A copy may be taller than its
	 * source (see `MorphTextPieceSplit`); a binding that clones the source's node
	 * gives the clone the copy's height.
	 */
	sources: Map<string, string>;
}

/** The id a piece copy is painted under. */
export function morphTextPieceId(sourceId: string, index: number): string {
	return `${sourceId}-morph-piece-${index}`;
}

function round(value: number): number {
	return Math.round(value * 1000) / 1000;
}

function clipOf(band: MorphTextBand): string {
	const points = morphTextBandOutline(band).map(([x, y]) => `${round(x)}px ${round(y)}px`);
	return `polygon(${points.join(', ')})`;
}

function transformOf(t: { dx: number; dy: number; sx: number; sy: number }): string {
	return `translate(${round(t.dx)}px, ${round(t.dy)}px) scale(${round(t.sx)}, ${round(t.sy)})`;
}

function keyframesOf(
	name: string,
	clip: string,
	from: { transform: string; opacity: number },
	to: { transform: string; opacity: number },
): string {
	return `
@keyframes ${name} {
\tfrom {
\t\ttransform: ${from.transform};
\t\tclip-path: ${clip};
\t\topacity: ${from.opacity};
\t}
\tto {
\t\ttransform: ${to.transform};
\t\tclip-path: ${clip};
\t\topacity: ${to.opacity};
\t}
}`;
}

/** Whether `pair` is a text pair the pieces should carry (and how), or `undefined`. */
function piecesFor(pair: MorphPair, ghostIds: ReadonlySet<string>, measure?: MorphTextMeasure) {
	const { fromElement: from, toElement: to } = pair;
	if (
		!ghostIds.has(from.id) ||
		morphTextReplacedInSlot(from, to) ||
		morphOffstageTextSizeChange(from, to) ||
		morphOffstageTextAdoptsTarget(from, to) ||
		resolveMorphTextAlignment(from, to, measure)
	) {
		return undefined;
	}
	return resolveMorphTextPieces(from, to, measure);
}

/**
 * Split every matched text pair whose layout changed into its pieces.
 *
 * Only pairs whose ghost is painted qualify, and only those the whole-box
 * text mapping (`resolveMorphTextAlignment`) cannot carry: a pair it carries
 * already moves exactly, and one it refuses for a reason other than its line
 * breaks is refused here too.
 *
 * @param pairs - The matched pairs.
 * @param ghostIds - Outgoing ids the overlay paints.
 * @param durationMs - The morph's duration.
 * @param measure - Text measurement; defaults to a canvas.
 */
export function buildMorphTextPiecePlan(
	pairs: readonly MorphPair[],
	ghostIds: ReadonlySet<string>,
	durationMs: number,
	measure?: MorphTextMeasure,
): MorphTextPiecePlan {
	const plan: MorphTextPiecePlan = {
		pairs: [],
		groups: [],
		outgoingAnimations: new Map(),
		incomingAnimations: new Map(),
		keyframes: [],
		sources: new Map(),
	};
	for (const pair of pairs) {
		const split = piecesFor(pair, ghostIds, measure);
		if (!split?.pieces.length) {
			continue;
		}
		const from = { ...pair.fromElement, height: split.fromHeight } as PptxElement;
		const to = { ...pair.toElement, height: split.toHeight } as PptxElement;
		plan.pairs.push(pair);
		const fromOpacity = from.opacity ?? 1;
		const toOpacity = to.opacity ?? 1;
		const still = 'translate(0px, 0px) scale(1, 1)';
		split.pieces.forEach((piece, index) => {
			const outgoingId = morphTextPieceId(from.id, index);
			const incomingId = morphTextPieceId(to.id, index);
			const safe = (id: string) => `pptx-morph-piece-${id.replace(/[^a-zA-Z0-9]/gu, '')}`;
			plan.keyframes.push(
				keyframesOf(
					safe(outgoingId),
					clipOf(piece.fromBand),
					{ transform: still, opacity: fromOpacity },
					{ transform: transformOf(morphTextGhostEnd(piece.alignment, from)), opacity: 0 },
				),
				keyframesOf(
					safe(incomingId),
					clipOf(piece.toBand),
					{ transform: transformOf(morphTextIncomingStart(piece.alignment, to)), opacity: 0 },
					{ transform: still, opacity: toOpacity },
				),
			);
			plan.outgoingAnimations.set(
				outgoingId,
				`${safe(outgoingId)} ${durationMs}ms ${MORPH_EASING} forwards`,
			);
			plan.incomingAnimations.set(
				incomingId,
				`${safe(incomingId)} ${durationMs}ms ${MORPH_EASING} forwards`,
			);
			plan.sources.set(outgoingId, pair.fromElement.id);
			plan.sources.set(incomingId, pair.toElement.id);
			plan.groups.push({
				outgoing: { ...from, id: outgoingId } as PptxElement,
				incoming: { ...to, id: incomingId } as PptxElement,
			});
		});
	}
	return plan;
}
