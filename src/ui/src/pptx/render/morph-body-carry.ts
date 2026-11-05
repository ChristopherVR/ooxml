/**
 * The geometry carry (see `morph-shape-carry`) on the BODY copies of a shape
 * whose text is carried apart from its body (see `morph-text-body`).
 *
 * Those copies travel on the plain box journey and dissolve into each other,
 * so each gets the carry in its own half's frame: both paint the same outline
 * and corners at every frame, and the dissolve between them is invisible
 * where the bodies agree. This is the case the measurement was first made on:
 * a rounded rectangle with a 3 pt outline and a label, growing 1.4x, whose
 * outline PowerPoint keeps at 3 pt with round corners throughout.
 *
 * @module render/morph-body-carry
 */
import { buildMorphCarryAnimations } from './morph-carry';
import { morphCopyId } from './morph-text-body';
import type { MorphTextPiecePlan } from './morph-text-piece-plan';
import type { MorphPair } from './morph-types';

/**
 * Add the carry of every split pair's body copies to the piece plan that
 * already animates them, and their outline tracks to `outlineAnimations`.
 */
export function addMorphBodyCopyCarry(
	pairs: readonly MorphPair[],
	durationMs: number,
	pieces: MorphTextPiecePlan,
	outlineAnimations: Map<string, string>,
	customGeometryAnimations?: Map<string, string>,
): void {
	pairs.forEach((pair, index) => {
		const halves = [
			{ id: morphCopyId(pair.toElement.id, 'body'), half: 'to', map: pieces.incomingAnimations },
			{
				id: morphCopyId(pair.fromElement.id, 'body'),
				half: 'from',
				map: pieces.outgoingAnimations,
			},
		] as const;
		for (const { id, half, map } of halves) {
			const existing = map.get(id);
			if (existing === undefined) {
				continue;
			}
			let composed = existing;
			for (const animation of buildMorphCarryAnimations(
				pair,
				{ id, half, effects: false },
				durationMs,
				`b${half}${index}`,
			)) {
				pieces.keyframes.push(animation.keyframes);
				if (animation.target === 'outline') {
					outlineAnimations.set(id, animation.animation);
				} else if (animation.target === 'custom') {
					customGeometryAnimations?.set(id, animation.animation);
				} else {
					composed = `${composed}, ${animation.animation}`;
				}
			}
			map.set(id, composed);
		}
	});
}
