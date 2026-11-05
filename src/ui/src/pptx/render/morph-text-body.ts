/**
 * A morphing SHAPE's text, carried apart from the shape's body.
 *
 * `morph-text-frame` carries a restyled text box line for line, but only a box
 * that paints nothing else: the overlay moves whole elements, and the text of
 * a filled or outlined shape is painted by the same element as its fill. A
 * matched shape whose box moves or resizes while its font or size changes was
 * therefore interpolated as one picture: its text was stretched with the box,
 * and the two halves' lines drifted apart into a double image.
 *
 * PowerPoint 16 does not morph it as one picture. Measured (`CreateVideo`,
 * 60 fps, 720p) on a rounded rectangle with a 3 pt white outline moving from
 * 300 x 120 pt to 420 x 180 pt while its "Lorem ipsum" changes Calibri 28 pt ->
 * Arial 40 pt:
 *
 *   - the body's box follows the morph curve (its left edge 110 -> 616 px) and
 *     its outline stays 3 pt and its corners round for every frame;
 *   - the text stays centred on the box (its centre within 0.5 px of the box's
 *     at every frame) and grows on the SAME curve (30% of the way: the text 41%
 *     of the way from 189 to 300 px wide, the box 42% from 394 to 554 px), one
 *     sharp copy per glyph whose faces dissolve where they differ (the "o" of
 *     "Lorem" shows both faces mid-morph), never two offset lines;
 *   - the brightest stroke of the text stays at 255 for every frame, the
 *     additive blend `morph-text-blend` describes.
 *
 * The same holds while the body does more than resize: with an outer shadow
 * (cast by the body alone), with corners whose adjustment changes, and with a
 * horizontal or vertical flip that changes between the slides (the body
 * flips edge-on, the text never mirrors or squashes: it travels upright).
 *
 * So the pair is painted as two crossfade groups of copies, the way
 * `morph-text-piece-plan` paints the pieces of a rewrapped text: a BODY group
 * (both halves without their text, interpolating the box as the pair always
 * has, its outline tween included, dissolving into each other) and above it a
 * TEXT group (both halves without their body, carried glyph over glyph on the
 * text mapping). The pair's own ghost is not painted and its stage copy is
 * held invisible.
 *
 * Only what the copies can paint faithfully is split: the same text, no effect
 * but a shadow under an opaque fill (a glow, say, would be painted by both
 * copies or by neither), and only where the text mapping differs from the
 * box's (a box that only moves already carries its text exactly). The outline
 * still scales with the box, as it did before the split.
 *
 * @module render/morph-text-body
 */
import type { PptxElement } from 'ooxml-core/pptx';

import { shortestRotationTarget } from './morph-animation';
import type { MorphCrossfadeGroup } from './morph-crossfade-group';
import { buildMorphFlipTurn, morphFlipTurns } from './morph-flip-turn';
import { generateGeometryMorphAnimation } from './morph-geometry-keyframes';
import type { MorphLiftContext } from './morph-text-blend';
import { resolveMorphLiftCrossers } from './morph-text-blend';
import {
	morphBodyCopy,
	morphBoxEnd,
	morphBoxStart,
	morphCopiesCannotSplitEffects,
	morphCopyKeyframes,
	morphCopyTransform,
	morphTextCopy,
	morphUprightText,
	sameMorphTransform,
} from './morph-text-body-copies';
import type { MorphTextMeasure, MorphTextTransform } from './morph-text-frame';
import {
	morphTextGhostEnd,
	morphTextGhostStart,
	morphTextIncomingEnd,
	morphTextIncomingStart,
	resolveMorphTextAlignment,
} from './morph-text-frame';
import { paintsBody } from './morph-text-frame-read';
import type { MorphTextPiecePlan } from './morph-text-piece-plan';
import { travelledBox } from './morph-overlay-order';
import {
	morphOffstageTextAdoptsTarget,
	morphOffstageTextSizeChange,
	morphTextReplacedInSlot,
} from './morph-text-slot';
import type { MorphPair } from './morph-types';
import { MORPH_EASING } from './morph-types';

/** Which part of its source element a copy paints. */
export type MorphCopyPart = 'body' | 'text';

/** The copies that carry the text of a shape apart from its body. */
export interface MorphTextBodyPlan extends MorphTextPiecePlan {
	/** Copy id -> the part of its source it paints. */
	parts: Map<string, MorphCopyPart>;
	/** Incoming ids that rise into the overlay so the groups stay under them. */
	rising: Set<string>;
}

/** The id a body or text copy is painted under. */
export function morphCopyId(sourceId: string, part: MorphCopyPart): string {
	return `${sourceId}-morph-${part}`;
}

/** Whether either half is flipped, so every frame has to state its flips. */
function flipped(from: PptxElement, to: PptxElement): boolean {
	return Boolean(from.flipHorizontal || from.flipVertical || to.flipHorizontal || to.flipVertical);
}

function splittable(pair: MorphPair, ghostIds: ReadonlySet<string>, measure?: MorphTextMeasure) {
	const { fromElement: from, toElement: to } = pair;
	if (
		!ghostIds.has(from.id) ||
		!(paintsBody(from) || paintsBody(to)) ||
		morphCopiesCannotSplitEffects(from) ||
		morphCopiesCannotSplitEffects(to) ||
		morphTextReplacedInSlot(from, to) ||
		morphOffstageTextSizeChange(from, to) ||
		morphOffstageTextAdoptsTarget(from, to)
	) {
		return undefined;
	}
	// A flipped shape's text is an unflipped shape's (turned half way round by
	// a vertical flip), so its text copy is read and painted that way (see
	// `morphUprightText`).
	const a = morphUprightText(from);
	const b = morphUprightText(to);
	const alignment = a && b && resolveMorphTextAlignment(a, b, measure, { withBody: true });
	if (
		!a ||
		!b ||
		!alignment ||
		(!flipped(from, to) &&
			sameMorphTransform(morphTextIncomingStart(alignment, to), morphBoxStart(from, to)))
	) {
		return undefined;
	}
	return { alignment, upright: { from: a, to: b } };
}

/** One copy's two frames. */
interface Leg {
	part: MorphCopyPart;
	outgoing: [string, string];
	incoming: [string, string];
	extra?: string | undefined;
	extraKeyframes?: string | undefined;
}

/**
 * Split every matched shape whose text should be carried apart from its body.
 *
 * @param pairs - The matched pairs (object mode only).
 * @param context - The slides and what the overlay already paints, to keep
 *   whatever the incoming slide stacks above a split pair on top of it.
 * @param durationMs - The morph's duration.
 * @param measure - Text measurement; defaults to a canvas.
 */
export function buildMorphTextBodyPlan(
	pairs: readonly MorphPair[],
	context: MorphLiftContext,
	durationMs: number,
	measure?: MorphTextMeasure,
): MorphTextBodyPlan {
	const plan: MorphTextBodyPlan = {
		pairs: [],
		groups: [],
		outgoingAnimations: new Map(),
		incomingAnimations: new Map(),
		keyframes: [],
		sources: new Map(),
		parts: new Map(),
		rising: new Set(),
	};
	for (const [pairIndex, pair] of pairs.entries()) {
		const split = splittable(pair, context.ghostIds, measure);
		const alignment = split?.alignment;
		const { fromElement: from, toElement: to } = pair;
		const rising =
			alignment &&
			resolveMorphLiftCrossers(
				{ ...context, overlaid: new Set([...context.overlaid, ...plan.rising]) },
				travelledBox(from, to),
				from,
				to,
			);
		if (!split || !alignment || !rising) {
			continue;
		}
		plan.pairs.push(pair);
		for (const id of rising) {
			plan.rising.add(id);
		}
		const fromRotation = from.rotation ?? 0;
		const toRotation = to.rotation ?? 0;
		// The copies turn the short way round, as the pair's own halves do.
		const ghostTurnsTo = shortestRotationTarget(fromRotation, toRotation);
		const incomingTurnsFrom = shortestRotationTarget(toRotation, fromRotation);
		// The text copies turn from text angle to text angle: a vertical flip
		// adds half a turn to its text (see `morph-text-flip`).
		const upright = split.upright;
		const fromText = upright.from.rotation ?? 0;
		const toText = upright.to.rotation ?? 0;
		const flips = flipped(from, to);
		const still: MorphTextTransform = { dx: 0, dy: 0, sx: 1, sy: 1 };
		// The body flips edge-on through zero; the text copies are painted
		// unflipped (see `morphTextCopy`), so they state no flips at all.
		const own = (element: PptxElement) => (flips ? element : undefined);
		const incomingBodyId = morphCopyId(to.id, 'body');
		const outline = generateGeometryMorphAnimation(
			{ fromElement: from, toElement: { ...to, id: incomingBodyId } as PptxElement },
			durationMs,
			pairIndex,
		);
		// A flip that changes turns the body over in 3D, as a shape without text
		// does (see `morph-flip-turn`); the text copies above it stay upright.
		const turns = morphFlipTurns(from, to);
		const outTurn = turns
			? buildMorphFlipTurn(from, from, to, fromRotation, ghostTurnsTo)
			: undefined;
		const inTurn = turns
			? buildMorphFlipTurn(to, from, to, incomingTurnsFrom, toRotation)
			: undefined;
		const legs: Leg[] = [
			{
				part: 'body',
				outgoing: [
					outTurn?.start ?? morphCopyTransform(still, fromRotation, own(from)),
					outTurn?.end ?? morphCopyTransform(morphBoxEnd(from, to), ghostTurnsTo, own(to)),
				],
				incoming: [
					inTurn?.start ??
						morphCopyTransform(morphBoxStart(from, to), incomingTurnsFrom, own(from)),
					inTurn?.end ?? morphCopyTransform(still, toRotation, own(to)),
				],
				extra: outline?.animation,
				extraKeyframes: outline?.keyframes,
			},
			{
				part: 'text',
				outgoing: [
					morphCopyTransform(morphTextGhostStart(alignment, from), fromText),
					morphCopyTransform(
						morphTextGhostEnd(alignment, from),
						shortestRotationTarget(fromText, toText),
					),
				],
				incoming: [
					morphCopyTransform(
						morphTextIncomingStart(alignment, to),
						shortestRotationTarget(toText, fromText),
					),
					morphCopyTransform(morphTextIncomingEnd(alignment, to), toText),
				],
			},
		];
		for (const { part, outgoing, incoming, extra, extraKeyframes } of legs) {
			const outgoingId = morphCopyId(from.id, part);
			const incomingId = morphCopyId(to.id, part);
			const safe = (id: string) => `pptx-morph-${part}-${id.replace(/[^a-zA-Z0-9]/gu, '')}`;
			plan.keyframes.push(
				morphCopyKeyframes(
					safe(outgoingId),
					{ transform: outgoing[0], opacity: from.opacity ?? 1 },
					{ transform: outgoing[1], opacity: 0 },
				),
				morphCopyKeyframes(
					safe(incomingId),
					{ transform: incoming[0], opacity: 0 },
					{ transform: incoming[1], opacity: to.opacity ?? 1 },
				),
				...(extraKeyframes ? [extraKeyframes] : []),
			);
			plan.outgoingAnimations.set(
				outgoingId,
				`${safe(outgoingId)} ${durationMs}ms ${MORPH_EASING} forwards`,
			);
			plan.incomingAnimations.set(
				incomingId,
				`${safe(incomingId)} ${durationMs}ms ${MORPH_EASING} forwards${extra ? `, ${extra}` : ''}`,
			);
			plan.sources.set(outgoingId, from.id);
			plan.sources.set(incomingId, to.id);
			plan.parts.set(outgoingId, part);
			plan.parts.set(incomingId, part);
			const body = part === 'body';
			const copy = body ? morphBodyCopy : morphTextCopy;
			plan.groups.push({
				outgoing: copy(body ? from : upright.from, outgoingId),
				incoming: copy(body ? to : upright.to, incomingId),
			} satisfies MorphCrossfadeGroup);
		}
	}
	return plan;
}

/**
 * Add `body`'s copies to the piece plan the overlay already paints copies
 * from, so both are hidden, stacked and painted by the same code.
 */
export function mergeMorphTextBodyPlan(into: MorphTextPiecePlan, body: MorphTextBodyPlan): void {
	into.pairs.push(...body.pairs);
	into.groups.push(...body.groups);
	into.keyframes.push(...body.keyframes);
	for (const [id, animation] of body.outgoingAnimations) {
		into.outgoingAnimations.set(id, animation);
	}
	for (const [id, animation] of body.incomingAnimations) {
		into.incomingAnimations.set(id, animation);
	}
	for (const [id, source] of body.sources) {
		into.sources.set(id, source);
	}
}
