/**
 * Which carried-text pairs (see `morph-text-frame`) the overlay should paint as
 * an additive cross-dissolve group.
 *
 * Once both halves of a restyled text are laid glyph over glyph, stacking them
 * with ordinary source-over blending dims every glyph they share: at the middle
 * of the dissolve the ink reaches only `1 - t + t^2` = 0.75 of full strength
 * (see `morph-crossfade-group`). Measured on a 32 pt white line changing
 * Calibri -> Arial: PowerPoint 16 keeps the brightest stroke at 255 for every
 * frame, while the stacked halves dipped to 193 mid-morph. A crossfade group
 * sums the two halves instead, which is PowerPoint's blend.
 *
 * A group is painted above everything else in the overlay, so a pair is only
 * moved there when that cannot change what covers what: nothing either slide
 * stacks above it may come near the ground it travels over.
 *
 * @module render/morph-text-blend
 */
import type { PptxElement } from 'ooxml-core/pptx';

import { boxesOverlap, travelledBox } from './morph-overlay-order';
import type { MorphTextMeasure } from './morph-text-frame';
import { resolveMorphTextAlignment } from './morph-text-frame';
import type { MorphPair } from './morph-types';

/**
 * Incoming ids of the carried-text pairs to lift into a crossfade group.
 *
 * @param outgoing - The outgoing slide, flattened, in document order.
 * @param incoming - The incoming slide, flattened, in document order.
 * @param pairs - The matched pairs.
 * @param ghostIds - Outgoing ids the overlay paints.
 * @param dissolvingInIds - Incoming ids whose half fades in over its ghost.
 * @param measure - Text measurement, as for `resolveMorphTextAlignment`.
 */
export function resolveMorphTextBlendLifts(
	outgoing: readonly PptxElement[],
	incoming: readonly PptxElement[],
	pairs: readonly MorphPair[],
	ghostIds: ReadonlySet<string>,
	dissolvingInIds: ReadonlySet<string>,
	measure?: MorphTextMeasure,
): Set<string> {
	const counterpartOfFrom = new Map(pairs.map((pair) => [pair.fromElement.id, pair.toElement]));
	const counterpartOfTo = new Map(pairs.map((pair) => [pair.toElement.id, pair.fromElement]));
	const lifts = new Set<string>();
	for (const { fromElement, toElement } of pairs) {
		if (
			!ghostIds.has(fromElement.id) ||
			!dissolvingInIds.has(toElement.id) ||
			!resolveMorphTextAlignment(fromElement, toElement, measure)
		) {
			continue;
		}
		const box = travelledBox(fromElement, toElement);
		const inIndex = incoming.findIndex((element) => element.id === toElement.id);
		const outIndex = outgoing.findIndex((element) => element.id === fromElement.id);
		const coveredOnStage = incoming
			.slice(inIndex + 1)
			.some((element) => boxesOverlap(box, travelledBox(element, counterpartOfTo.get(element.id))));
		const coveredByGhost = outgoing
			.slice(outIndex + 1)
			.some(
				(element) =>
					ghostIds.has(element.id) &&
					boxesOverlap(box, travelledBox(element, counterpartOfFrom.get(element.id))),
			);
		if (inIndex >= 0 && outIndex >= 0 && !coveredOnStage && !coveredByGhost) {
			lifts.add(toElement.id);
		}
	}
	return lifts;
}
