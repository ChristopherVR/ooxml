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
 * PowerPoint keeps that blend when something stacked ABOVE the text crosses
 * its path. Measured (`CreateVideo`, 720p) on a 36 pt line changing Calibri ->
 * Arial while it slides under an opaque grey bar the slide stacks above it:
 * the bar covers the text on every frame, and the text left uncovered keeps
 * its brightest stroke at 255 (its 99th percentile too) through the whole
 * dissolve, where the stacked halves dipped to 205.
 *
 * A group is ranked at its incoming half's layer in the overlay's one stack
 * (`morph-overlay-stack`), where the pair's ghost was already painted, so
 * lifting a pair only changes what covers its INCOMING half: that half leaves
 * the live stage, under the whole overlay, for the pair's own layer. Whatever
 * the incoming slide stacks above the pair and paints in the overlay already
 * stays above it. What it paints only on the live stage would end up below
 * it, so each such shape that crosses the pair's path rises into the overlay
 * with it (and, in turn, whatever crosses that shape's path); when one cannot
 * rise, the pair stays where it was.
 *
 * @module render/morph-text-blend
 */
import type { PptxElement } from 'ooxml-core/pptx';

import { morphPairIncomingFadesIn, morphPairNeedsCrossfade } from './morph-animation';
import type { MorphBox } from './morph-overlay-order';
import { boxesOverlap, travelledBox } from './morph-overlay-order';
import type { MorphTextMeasure } from './morph-text-frame';
import { resolveMorphTextAlignment } from './morph-text-frame';
import type { MorphPair } from './morph-types';

/** What decides whether a shape crossing a lifted pair can rise with it. */
export interface MorphLiftContext {
	/** The outgoing slide, flattened, in document order. */
	outgoing: readonly PptxElement[];
	/** The incoming slide, flattened, in document order. */
	incoming: readonly PptxElement[];
	/** The matched pairs. */
	pairs: readonly MorphPair[];
	/** Outgoing ids the overlay paints. */
	ghostIds: ReadonlySet<string>;
	/** Incoming ids the overlay already paints (lifted, or carried by copies). */
	overlaid: ReadonlySet<string>;
	/** Whether an incoming element has an animation the overlay can run instead. */
	animated: (id: string) => boolean;
}

/**
 * The incoming ids that have to rise into the overlay so a pair lifted over
 * `box` keeps everything the incoming slide stacks above it on top, or
 * `undefined` when one of them cannot.
 *
 * @param context - The slides, pairs and what the overlay already paints.
 * @param box - The ground the lifted pair travels over.
 * @param from - The pair's outgoing half.
 * @param to - The pair's incoming half.
 */
export function resolveMorphLiftCrossers(
	context: MorphLiftContext,
	box: MorphBox,
	from: PptxElement,
	to: PptxElement,
): Set<string> | undefined {
	const { outgoing, incoming, pairs, ghostIds, overlaid, animated } = context;
	const counterpartOfFrom = new Map(pairs.map((pair) => [pair.fromElement.id, pair.toElement]));
	const counterpartOfTo = new Map(pairs.map((pair) => [pair.toElement.id, pair.fromElement]));
	const outIndex = outgoing.findIndex((element) => element.id === from.id);
	const inIndex = incoming.findIndex((element) => element.id === to.id);
	if (outIndex < 0 || inIndex < 0) {
		return undefined;
	}
	// A departing shape has no layer on the incoming slide to rank it by, so
	// one that the outgoing slide stacks above the pair keeps the pair down.
	const departingAbove = outgoing
		.slice(outIndex + 1)
		.some(
			(element) =>
				ghostIds.has(element.id) &&
				!counterpartOfFrom.has(element.id) &&
				boxesOverlap(box, travelledBox(element)),
		);
	if (departingAbove) {
		return undefined;
	}
	const rising = new Set<string>();
	const pending: { box: MorphBox; index: number }[] = [{ box, index: inIndex }];
	while (pending.length > 0) {
		const next = pending.pop()!;
		for (let index = next.index + 1; index < incoming.length; index++) {
			const element = incoming[index]!;
			const counterpart = counterpartOfTo.get(element.id);
			const path = travelledBox(element, counterpart);
			if (rising.has(element.id) || overlaid.has(element.id) || !boxesOverlap(next.box, path)) {
				continue;
			}
			const ghosted = counterpart !== undefined && ghostIds.has(counterpart.id);
			// A ghost that keeps its opacity is painted in the overlay at this
			// shape's layer and lands exactly on it: it already covers the pair.
			if (ghosted && !morphPairNeedsCrossfade(counterpart, element)) {
				continue;
			}
			// One whose ghost dissolves over a half held at full strength cannot
			// rise without its two halves summing past full strength.
			const canRise =
				animated(element.id) && (!ghosted || morphPairIncomingFadesIn(counterpart, element, true));
			if (!canRise) {
				return undefined;
			}
			rising.add(element.id);
			pending.push({ box: path, index });
		}
	}
	return rising;
}

/**
 * Incoming ids of the carried-text pairs to lift into a crossfade group, plus
 * the ids of the shapes that rise into the overlay with them.
 *
 * @param context - The slides, pairs and what the overlay already paints.
 * @param dissolvingInIds - Incoming ids whose half fades in over its ghost.
 * @param measure - Text measurement, as for `resolveMorphTextAlignment`.
 */
export function resolveMorphTextBlendLifts(
	context: MorphLiftContext,
	dissolvingInIds: ReadonlySet<string>,
	measure?: MorphTextMeasure,
): Set<string> {
	const lifts = new Set<string>();
	for (const { fromElement, toElement } of context.pairs) {
		if (
			!context.ghostIds.has(fromElement.id) ||
			!dissolvingInIds.has(toElement.id) ||
			!resolveMorphTextAlignment(fromElement, toElement, measure)
		) {
			continue;
		}
		const rising = resolveMorphLiftCrossers(
			{ ...context, overlaid: new Set([...context.overlaid, ...lifts]) },
			travelledBox(fromElement, toElement),
			fromElement,
			toElement,
		);
		if (rising) {
			lifts.add(toElement.id);
			for (const id of rising) {
				lifts.add(id);
			}
		}
	}
	return lifts;
}
