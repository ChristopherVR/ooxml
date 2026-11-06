/**
 * The other end of a morph, handed to the silhouette filter of the shape that
 * arrives (see `morph-effect-silhouette`).
 *
 * A filter is built from one element, and a morphing shape paints the effects of
 * BOTH its ends, so the plan registers the outgoing end here, scaled into the
 * arriving shape's frame, and `getSilhouetteEffects` adds its layers to the
 * arriving shape's filter. The entry matches by id AND by the arriving shape's
 * own effects, so a stale entry never touches another shape; the layers it adds
 * sit at opacity `1 - p` and read `p` from the container's keyframes, which end
 * at `p = 1`, so a shape that renders after its morph looks exactly as it did.
 *
 * @module render/silhouette-morph-registry
 */
import type { SilhouetteEffectParams } from './silhouette-effects';

/** The outgoing end of a morph, with what rebuilds the arriving shape's filter. */
export interface SilhouetteMorphEntry {
	/** The arriving shape's own effects. */
	readonly params: SilhouetteEffectParams;
	readonly width: number;
	readonly height: number;
	/** The arriving shape's own effects, serialized (the match key). */
	readonly own: string;
	/** The outgoing end's effects, in the arriving shape's frame. */
	readonly out: SilhouetteEffectParams;
}

const entries = new Map<string, SilhouetteMorphEntry>();

/** Registers the outgoing end of the morph that brings `elementId` in. */
export function registerSilhouetteMorph(
	elementId: string,
	own: SilhouetteEffectParams,
	out: SilhouetteEffectParams,
	size: { width: number; height: number },
): void {
	entries.set(elementId, { params: own, ...size, own: JSON.stringify(own), out });
}

/** The registered morph that brings `elementId` in, if any. */
export function silhouetteMorphEntry(elementId: string): SilhouetteMorphEntry | undefined {
	return entries.get(elementId);
}

/** The outgoing end's effects for a shape with `own` effects, if a morph brings it in. */
export function silhouetteMorphFor(
	elementId: string,
	own: SilhouetteEffectParams,
): SilhouetteEffectParams | undefined {
	const entry = entries.get(elementId);
	return entry && entry.own === JSON.stringify(own) ? entry.out : undefined;
}
