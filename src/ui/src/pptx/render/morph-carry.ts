/**
 * The geometry and effect CARRY animations of a morph's matched shapes: what
 * `morph-shape-carry` and `morph-effect-carry` describe, turned into the
 * animations the plan routes onto each half.
 *
 * A carry only means anything on a half whose journey is the plain box
 * journey (`translate` + `scale` from one box onto the other): it is written in
 * that box's local frame. A text box carried line for line, an in-slot rewrite
 * or an off-stage single copy travels some other way and is left alone.
 *
 * @module render/morph-carry
 */
import type { PptxElement } from 'ooxml-core/pptx';

import type { MorphEffectPainter } from './morph-effect-carry';
import { buildMorphEffectCarry } from './morph-effect-carry';
import type { MorphCarryHalf } from './morph-shape-carry';
import { buildMorphShapeCarry } from './morph-shape-carry';
import type { MorphTextMeasure } from './morph-text-frame';
import { resolveMorphTextAlignment } from './morph-text-frame';
import { morphOffstageTextSizeChange, morphTextReplacedInSlot } from './morph-text-slot';
import type { MorphAnimationStyle, MorphMode, MorphPair } from './morph-types';
import { MORPH_EASING } from './morph-types';

/** Whether both halves of a pair travel on the box journey. */
export function morphPairRidesBox(
	from: PptxElement,
	to: PptxElement,
	mode: MorphMode,
	measure?: MorphTextMeasure,
): boolean {
	if (morphTextReplacedInSlot(from, to) || morphOffstageTextSizeChange(from, to)) {
		return false;
	}
	return mode !== 'object' || resolveMorphTextAlignment(from, to, measure) === undefined;
}

/** Which effects a half paints itself, rather than a ghost dissolving them over it. */
export interface MorphCarryTarget {
	/** The element id the half is painted under. */
	readonly id: string;
	readonly half: MorphCarryHalf;
	/** Whether this half is the only painter of the pair's effects. */
	readonly effects: boolean;
	/**
	 * `underGhost` when the outgoing ghost dissolves over this half and paints
	 * the outgoing effects itself (see `morph-effect-carry`). Defaults to `alone`.
	 */
	readonly effectPainter?: MorphEffectPainter;
}

function safe(id: string): string {
	return id.replace(/[^a-zA-Z0-9]/gu, '');
}

/**
 * The carry animations of one half: on its container (radii, effects, clip)
 * and on its outline path (`target: 'outline'`).
 */
export function buildMorphCarryAnimations(
	pair: MorphPair,
	target: MorphCarryTarget,
	durationMs: number,
	tag: string,
): MorphAnimationStyle[] {
	const { fromElement: from, toElement: to } = pair;
	const shape = buildMorphShapeCarry(from, to, target.half);
	const effects =
		target.effects && from.type === 'shape' && to.type === 'shape'
			? buildMorphEffectCarry(from, to, target.half, target.effectPainter)
			: undefined;
	const start = [...(shape?.startDeclarations ?? [])];
	const end = [...(shape?.endDeclarations ?? [])];
	const base = `pptx-morph-carry-${tag}-${safe(target.id)}`;
	const animations: MorphAnimationStyle[] = [];
	if (start.length > 0) {
		const lines = (list: string[]) => list.map((declaration) => `\t\t${declaration};`).join('\n');
		animations.push({
			elementId: target.id,
			animation: `${base} ${durationMs}ms ${MORPH_EASING} forwards`,
			keyframes: `\n@keyframes ${base} {\n\tfrom {\n${lines(start)}\n\t}\n\tto {\n${lines(end)}\n\t}\n}`,
		});
	}
	if (effects) {
		animations.push({
			elementId: target.id,
			animation: `${base}-fx ${durationMs}ms linear forwards`,
			keyframes: `
@keyframes ${base}-fx {
${effects.keyframes}
}`,
		});
	}
	if (shape?.clipKeyframes) {
		animations.push({
			elementId: target.id,
			animation: `${base}-clip ${durationMs}ms linear forwards`,
			keyframes: `\n@keyframes ${base}-clip {\n${shape.clipKeyframes}\n}`,
		});
	}
	if (shape?.outlineKeyframes) {
		animations.push({
			elementId: target.id,
			animation: `${base}-outline ${durationMs}ms linear forwards`,
			keyframes: `\n@keyframes ${base}-outline {\n${shape.outlineKeyframes}\n}`,
			target: 'outline',
		});
	}
	return animations;
}

/** Result of {@link generateMorphCarryAnimations}. */
export interface MorphCarryResult {
	readonly animations: MorphAnimationStyle[];
	/** Pairs whose incoming half got a geometry carry (their outline needs no other tween). */
	readonly carried: ReadonlySet<MorphPair>;
}

/**
 * Carry animations for every matched pair whose halves ride the box journey.
 *
 * @param pairs - The matched pairs.
 * @param durationMs - The morph's duration.
 * @param ghostIds - Outgoing ids the overlay paints (see `resolveMorphGhostIds`).
 * @param needsCrossfade - Whether a pair's ghost dissolves over its incoming half.
 * @param inert - Whether a pair is inert (it is then not animated at all).
 * @param mode - Morph granularity.
 * @param measure - Text measurement for the text-alignment check.
 * @param incomingFadesIn - Whether a ghosted pair's incoming half fades in (see
 *   `morphPairIncomingFadesIn`). When it does not, the ghost dissolves over a
 *   full-strength half, which then fades the incoming effects in under it.
 *   Without it such a half carries no effects, as before.
 */
export function generateMorphCarryAnimations(
	pairs: readonly MorphPair[],
	durationMs: number,
	ghostIds: ReadonlySet<string>,
	needsCrossfade: (from: PptxElement, to: PptxElement) => boolean,
	inert: (from: PptxElement, to: PptxElement) => boolean,
	mode: MorphMode,
	measure?: MorphTextMeasure,
	incomingFadesIn?: (from: PptxElement, to: PptxElement, ghosted: boolean) => boolean,
): MorphCarryResult {
	const animations: MorphAnimationStyle[] = [];
	const carried = new Set<MorphPair>();
	pairs.forEach((pair, index) => {
		const { fromElement: from, toElement: to } = pair;
		if (
			from.type !== 'shape' ||
			to.type !== 'shape' ||
			inert(from, to) ||
			!morphPairRidesBox(from, to, mode, measure)
		) {
			return;
		}
		const ghosted = ghostIds.has(from.id);
		const fades = needsCrossfade(from, to);
		const underGhost =
			ghosted && fades && incomingFadesIn !== undefined && !incomingFadesIn(from, to, true);
		const incoming = buildMorphCarryAnimations(
			pair,
			{
				id: to.id,
				half: 'to',
				effects: !(ghosted && fades) || underGhost,
				...(underGhost ? { effectPainter: 'underGhost' as const } : {}),
			},
			durationMs,
			String(index),
		);
		if (buildMorphShapeCarry(from, to, 'to')) {
			carried.add(pair);
		}
		animations.push(...incoming);
		if (ghosted) {
			animations.push(
				...buildMorphCarryAnimations(
					pair,
					{ id: from.id, half: 'from', effects: !fades },
					durationMs,
					`g${index}`,
				),
			);
		}
	});
	return { animations, carried };
}
