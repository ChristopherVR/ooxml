/**
 * A soft edge (`a:softEdge`) dissolved the way PowerPoint dissolves it.
 *
 * Measured on PowerPoint 16 (`CreateVideo`, 60 fps, 720p, Morph by object) on a
 * rectangle with a 20 pt soft edge growing from 200 x 120 pt to 440 x 240 pt:
 * mid-morph there are TWO renders of the shape, each with its own feather
 * scaled with the box (`W(p) / W0` and `W(p) / W1` times the authored radius),
 * and they are cross-dissolved per pixel as `(1 - p) A + p B`. That blend
 * reproduces the feather profile across the edge within 0.002-0.005 coverage at
 * every frame; one feather riding the box is 0.01-0.07 off.
 *
 * The soft edge is an SVG filter fixed in the element's own frame, which no
 * keyframe can reach, so the feather cannot be tweened on one element. But a
 * ghost (the outgoing half, scaled up from its own size) and the incoming half
 * (scaled down from its final size) each carry their own feather in exactly the
 * frame PowerPoint scales it in. Dissolving the two additively is the measured
 * blend, so a soft-edged pair is handed to the crossfade group like restyled
 * text is (see `morph-text-blend` and `morph-crossfade-group`).
 *
 * @module render/morph-soft-edge-blend
 */
import type { PptxElement } from 'ooxml-core/pptx';

import { travelledBox } from './morph-overlay-order';
import type { MorphLiftContext } from './morph-text-blend';
import { resolveMorphLiftCrossers } from './morph-text-blend';

/** Sub-pixel tolerance for calling a pair's size unchanged. */
const SIZE_EPSILON = 0.5;

function softEdgeRadius(element: PptxElement): number {
	const style = (element as { shapeStyle?: { softEdgeRadius?: number } }).shapeStyle;
	return typeof style?.softEdgeRadius === 'number' && style.softEdgeRadius > 0
		? style.softEdgeRadius
		: 0;
}

/**
 * Whether a matched pair's soft edge looks different at its two ends: the
 * radius differs, or it is the same radius on a box of another size (the
 * feather scales with the box, so each end is a different render).
 */
export function morphSoftEdgeBlends(from: PptxElement, to: PptxElement): boolean {
	if (from.type !== 'shape' || to.type !== 'shape') {
		return false;
	}
	const a = softEdgeRadius(from);
	const b = softEdgeRadius(to);
	if (a === 0 && b === 0) {
		return false;
	}
	return (
		a !== b ||
		Math.abs(from.width - to.width) > SIZE_EPSILON ||
		Math.abs(from.height - to.height) > SIZE_EPSILON
	);
}

/**
 * Incoming ids of the soft-edged pairs to lift into a crossfade group, plus the
 * shapes that rise into the overlay with them (the same rule as restyled text).
 *
 * @param context - The slides, pairs and what the overlay already paints.
 * @param dissolvingInIds - Incoming ids whose half fades in over its ghost.
 */
export function resolveMorphSoftEdgeLifts(
	context: MorphLiftContext,
	dissolvingInIds: ReadonlySet<string>,
): Set<string> {
	const lifts = new Set<string>();
	for (const { fromElement, toElement } of context.pairs) {
		if (
			!context.ghostIds.has(fromElement.id) ||
			!dissolvingInIds.has(toElement.id) ||
			!morphSoftEdgeBlends(fromElement, toElement)
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
