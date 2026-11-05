/**
 * Which matched pairs have their text carried apart from the whole element,
 * and the copies that carry it: piece by piece, glyph by glyph, or body apart
 * from text (see `morph-text-piece-plan`, `morph-text-glyph-plan`,
 * `morph-text-body`). One dispatch for the three, in the order they claim
 * pairs, so no pair is carried twice.
 *
 * @module render/morph-plan-carry
 */
import type { PptxElement } from 'ooxml-core/pptx';

import { addMorphBodyCopyCarry } from './morph-body-carry';
import type { MorphTextBodyPlan } from './morph-text-body';
import { buildMorphTextBodyPlan, mergeMorphTextBodyPlan } from './morph-text-body';
import type { MorphTextMeasure } from './morph-text-frame';
import { buildMorphTextGlyphPlan } from './morph-text-glyph-plan';
import type { MorphTextPiecePlan } from './morph-text-piece-plan';
import { buildMorphTextPiecePlan } from './morph-text-piece-plan';
import { morphOffstageTextAdoptsTarget } from './morph-text-slot';
import type { MorphMode, MorphPair } from './morph-types';

/** Everything {@link planMorphTextCarry} reads about the transition. */
export interface MorphTextCarryInput {
	readonly mode: MorphMode;
	readonly durationMs: number;
	readonly measureText: MorphTextMeasure | undefined;
	readonly pairs: readonly MorphPair[];
	readonly ghostIds: ReadonlySet<string>;
	readonly flattenedOutgoing: readonly PptxElement[];
	readonly flattenedIncoming: readonly PptxElement[];
	/** Whether the plan holds a container animation for an incoming id yet. */
	readonly animated: (id: string) => boolean;
	readonly outlineAnimations: Map<string, string>;
	readonly customGeometryAnimations: Map<string, string>;
	readonly outlineStrandAnimations: Map<string, string[]>;
}

/** The copies that carry text apart from its element. */
export interface MorphTextCarry {
	/** Outgoing text id -> the incoming text it adopts the layout of while offstage. */
	offstageRestyles: Map<string, PptxElement>;
	/** Piece, glyph and body copies merged into one plan. */
	pieces: MorphTextPiecePlan;
	/** The body-apart-from-text half of {@link MorphTextCarry.pieces}. */
	bodies: MorphTextBodyPlan;
}

/** Plan every text carry of the transition and merge them into one piece plan. */
export function planMorphTextCarry(input: MorphTextCarryInput): MorphTextCarry {
	const { mode, durationMs, measureText, pairs, ghostIds } = input;
	const offstageRestyles = new Map(
		pairs
			.filter(({ fromElement, toElement }) => morphOffstageTextAdoptsTarget(fromElement, toElement))
			.map(({ fromElement, toElement }) => [fromElement.id, toElement] as const),
	);
	// A text pair whose line breaks change is carried piece by piece; its own
	// ghost and stage copy give way to the pieces' copies (see
	// `morph-text-piece-plan`).
	const pieces = buildMorphTextPiecePlan(
		mode === 'object' ? pairs : [],
		ghostIds,
		durationMs,
		measureText,
	);
	// A text box that turns while its type changes is carried glyph by glyph
	// (see `morph-text-glyph-plan`), through the same copies.
	const glyphs = buildMorphTextGlyphPlan(
		mode === 'object' ? pairs.filter((pair) => !pieces.pairs.includes(pair)) : [],
		ghostIds,
		durationMs,
		measureText,
	);
	mergeMorphTextBodyPlan(pieces, glyphs);
	// A shape whose text changes face or size while its box moves carries the
	// text apart from its body (see `morph-text-body`), through the same copies.
	const bodies = buildMorphTextBodyPlan(
		mode === 'object' ? pairs.filter((pair) => !pieces.pairs.includes(pair)) : [],
		{
			outgoing: input.flattenedOutgoing,
			incoming: input.flattenedIncoming,
			pairs,
			ghostIds,
			overlaid: new Set(pieces.pairs.map((pair) => pair.toElement.id)),
			animated: input.animated,
		},
		durationMs,
		measureText,
	);
	mergeMorphTextBodyPlan(pieces, bodies);
	addMorphBodyCopyCarry(
		bodies.pairs,
		durationMs,
		pieces,
		input.outlineAnimations,
		input.customGeometryAnimations,
		input.outlineStrandAnimations,
	);
	return { offstageRestyles, pieces, bodies };
}
