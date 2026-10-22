/**
 * The stacking order of everything a morph's transition overlay paints.
 *
 * The overlay paints ghosts of the outgoing slide and, above holding ghosts
 * that would hide them, lifted copies of incoming shapes (see
 * `morph-overlay-order`). It used to paint them as two flat layers, every ghost
 * below every lifted arrival, which breaks as soon as a ghost belongs ABOVE an
 * arrival. That is the ordinary case once a backdrop pair is ghosted:
 * everything arriving over the backdrop is lifted, including a full-slide
 * picture fading in UNDER the title, and the title's travelling ghost then
 * dimmed beneath that picture for the whole morph and snapped bright when the
 * overlay came down (issue #141).
 *
 * This module ranks both kinds in one order and hands the bindings runs of
 * consecutive same-slide shapes, so every binding paints the same stack: it
 * renders each run as one layer, back to front, with that slide as context.
 *
 * @module render/morph-overlay-layers
 */
import type { PptxElement } from 'ooxml-core/pptx';

import type { MorphBox } from './morph-overlay-order';
import { boxesOverlap, travelledBox } from './morph-overlay-order';
import type { MorphPair } from './morph-types';

/**
 * One run of the transition overlay: consecutive shapes painted from the same
 * slide, so a binding can render the run with that slide as its context.
 */
export interface MorphOverlayLayer {
	/**
	 * `'outgoing'` for ghosts (rendered against the outgoing slide and scoped by
	 * the outgoing rules), `'incoming'` for lifted arrivals (rendered against
	 * the incoming slide and scoped by the `lifted` rules).
	 */
	readonly side: 'outgoing' | 'incoming';
	/** The run's shapes, back to front. */
	readonly elements: readonly PptxElement[];
}

/**
 * Rank every shape of both slides back to front, anchoring each matched pair
 * to the layer the INCOMING slide gives it.
 *
 * Every incoming shape ranks at its own place in the incoming slide, and a
 * matched outgoing shape shares its counterpart's rank, so a pair sits where
 * it will rest once the overlay comes down. A departure keeps its place
 * between the matched shapes around it on the outgoing slide, as low as that
 * allows: just above the matched shape below it. When those two neighbours
 * swapped places (the lower one now ranks above the upper one), it goes just
 * below the upper one instead, so whatever covered it when the morph began
 * still covers it. Ranking it after the lower neighbour alone let a departing
 * full-slide picture that the backdrop hid jump above the backdrop when that
 * neighbour moved far up the incoming stack (issue #141's slides 9 to 10).
 *
 * @param outgoing - The outgoing slide's elements, flattened, in document order.
 * @param incoming - The incoming slide's elements, flattened, in document order.
 * @param pairs - The matched pairs.
 * @returns Element id -> rank; higher is nearer the viewer.
 */
export function buildMorphIncomingAnchoredOrder(
	outgoing: readonly PptxElement[],
	incoming: readonly PptxElement[],
	pairs: readonly MorphPair[],
): Map<string, number> {
	const rank = new Map<string, number>(incoming.map((element, index) => [element.id, index]));
	const partnerOf = new Map(pairs.map((pair) => [pair.fromElement.id, pair.toElement.id]));
	let run: string[] = [];
	let floor = -1;
	/**
	 * Slot the pending departures between the matched shapes around them:
	 * just above `floor` when it is below `ceiling`, else (the two pairs crossed
	 * over) just below `ceiling`.
	 */
	const flush = (ceiling: number): void => {
		const base = floor < ceiling ? floor : ceiling - 1;
		const span = floor < ceiling ? Math.min(1, ceiling - floor) : 1;
		run.forEach((id, index) => rank.set(id, base + (span * (index + 1)) / (run.length + 1)));
		run = [];
	};
	for (const element of outgoing) {
		const partner = partnerOf.get(element.id);
		const partnerRank = partner === undefined ? undefined : rank.get(partner);
		if (partnerRank === undefined) {
			run.push(element.id);
			continue;
		}
		flush(partnerRank);
		rank.set(element.id, partnerRank);
		floor = partnerRank;
	}
	flush(incoming.length);
	return rank;
}

/**
 * Ghost every stage-journeyed pair a holding ghost would hide although the
 * incoming slide stacks the pair ABOVE that ghost's counterpart.
 *
 * A pair the overlay does not ghost travels on the live stage, and the whole
 * overlay sits above the stage. That is only harmless while no ghost that
 * stays opaque for the morph covers the pair's path, or while the pair really
 * belongs under it. Issue #141's slides 9 to 10 break it: the leftmost column
 * picture is first in the outgoing slide's order, so nothing painted sits
 * below it and it was left to the stage, but the incoming slide stacks it over
 * the full-slide backdrop, whose unchanged ghost then hid the column for its
 * whole climb into place; it appeared in one frame when the overlay came down.
 * Painting such a pair's ghost puts it in the overlay at its incoming layer
 * (see {@link resolveMorphOverlayLayers}), above the backdrop.
 *
 * Runs to a fixed point, since a pair ghosted here holds too and can hide
 * another.
 *
 * @param ghostIds - The outgoing ids already painted; extended in place.
 * @param outgoing - The outgoing slide's elements, flattened, in document order.
 * @param incoming - The incoming slide's elements, flattened, in document order.
 * @param pairs - The matched pairs.
 * @param holds - Whether a pair's ghost keeps its opacity for the whole morph
 *   (its appearance did not change).
 * @returns `ghostIds`, for chaining.
 */
export function ghostMorphPairsHiddenByGhosts(
	ghostIds: Set<string>,
	outgoing: readonly PptxElement[],
	incoming: readonly PptxElement[],
	pairs: readonly MorphPair[],
	holds: (pair: MorphPair) => boolean,
): Set<string> {
	const rank = buildMorphIncomingAnchoredOrder(outgoing, incoming, pairs);
	const boxOf = (pair: MorphPair): MorphBox => travelledBox(pair.fromElement, pair.toElement);
	let changed = true;
	while (changed) {
		changed = false;
		const shields = pairs.filter((pair) => ghostIds.has(pair.fromElement.id) && holds(pair));
		for (const pair of pairs) {
			if (ghostIds.has(pair.fromElement.id)) {
				continue;
			}
			const mine = rank.get(pair.fromElement.id) ?? 0;
			const box = boxOf(pair);
			if (
				shields.some(
					(shield) =>
						(rank.get(shield.fromElement.id) ?? 0) < mine && boxesOverlap(boxOf(shield), box),
				)
			) {
				ghostIds.add(pair.fromElement.id);
				changed = true;
			}
		}
	}
	return ghostIds;
}

/**
 * The order the overlay paints its ghosts and lifted arrivals in, as runs a
 * binding maps one-to-one onto stacked layers.
 *
 * Everything is ranked in one order anchored to the incoming slide
 * ({@link buildMorphIncomingAnchoredOrder}), because a travelling ghost IS its
 * counterpart mid-flight and the overlay has to end the morph stacked exactly
 * like the incoming slide, or something changes layer in the frame the overlay
 * is torn down. That also governs ghosts among themselves: painted in the
 * OUTGOING slide's order, a photo sliding in from off-slide (last in that
 * order) covered the title travelling across it although the incoming slide
 * stacks the title on top (issue #141, slides 3 to 4).
 *
 * A ghost and its own lifted counterpart share a rank; the arrival is painted
 * over its ghost, the way it dissolves in.
 *
 * @param ghosts - The outgoing elements the overlay paints.
 * @param lifted - The incoming elements the overlay paints.
 * @param outgoing - The outgoing slide's elements, flattened, in document order.
 * @param incoming - The incoming slide's elements, flattened, in document order.
 * @param pairs - The matched pairs.
 * @returns The runs, back to front; empty when the overlay paints nothing.
 */
export function resolveMorphOverlayLayers(
	ghosts: readonly PptxElement[],
	lifted: readonly PptxElement[],
	outgoing: readonly PptxElement[],
	incoming: readonly PptxElement[],
	pairs: readonly MorphPair[],
): MorphOverlayLayer[] {
	const rank = buildMorphIncomingAnchoredOrder(outgoing, incoming, pairs);
	const entries = [
		...ghosts.map((element, index) => ({ element, side: 'outgoing' as const, index })),
		...lifted.map((element, index) => ({ element, side: 'incoming' as const, index })),
	];
	entries.sort(
		(a, b) =>
			(rank.get(a.element.id) ?? 0) - (rank.get(b.element.id) ?? 0) ||
			Number(a.side === 'incoming') - Number(b.side === 'incoming') ||
			a.index - b.index,
	);
	const layers: { side: 'outgoing' | 'incoming'; elements: PptxElement[] }[] = [];
	for (const entry of entries) {
		const last = layers.at(-1);
		if (last?.side === entry.side) {
			last.elements.push(entry.element);
		} else {
			layers.push({ side: entry.side, elements: [entry.element] });
		}
	}
	return layers;
}
