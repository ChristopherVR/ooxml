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
 * So the pair is painted as two crossfade groups of copies, the way
 * `morph-text-piece-plan` paints the pieces of a rewrapped text: a BODY group
 * (both halves without their text, interpolating the box as the pair always
 * has, dissolving into each other) and above it a TEXT group (both halves
 * without their body, carried glyph over glyph on the text mapping). The
 * pair's own ghost is not painted and its stage copy is held invisible.
 *
 * Only what the copies can paint faithfully is split: the same text in the
 * same geometry and adjustments, no effects (a shadow or glow would be painted
 * by both copies, or by neither), and only where the text mapping differs from
 * the box's (a box that only moves already carries its text exactly). The
 * outline still scales with the box, as it did before the split.
 *
 * @module render/morph-text-body
 */
import type { PptxElement } from 'ooxml-core/pptx';
import { hasShapeProperties } from 'ooxml-core/pptx';

import type { MorphCrossfadeGroup } from './morph-crossfade-group';
import { serializeShapeAdjustments } from './morph-geometry-keyframes';
import type { MorphLiftContext } from './morph-text-blend';
import { resolveMorphLiftCrossers } from './morph-text-blend';
import type { MorphTextMeasure, MorphTextTransform } from './morph-text-frame';
import {
	morphTextGhostEnd,
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

/** Style keys of an effect a copy could not paint once and only once. */
const EFFECT_KEYS =
	/^(shadow|innerShadow|glow|softEdge|reflection|effectList|effectDag|outerShadow|presetShadow|blur|scene3d|shape3d|dag|lineShadow|lineGlow|shapeFillOverlay|fillOverlay)/u;
/** Style keys that paint the body: fills, outlines and their theme references. */
const BODY_KEYS =
	/^(fill|stroke|useBackgroundFill|inheritedFillStyle|inheritedLineStyle|lnRef|compoundLine|lineAlignment|customDash)/u;

function hasEffects(element: PptxElement): boolean {
	const style = hasShapeProperties(element) ? element.shapeStyle : undefined;
	const inherited = style?.inheritedEffectStyle as object | undefined;
	return (
		Object.keys(inherited ?? {}).length > 0 ||
		Object.entries(style ?? {}).some(
			([key, value]) => EFFECT_KEYS.test(key) && value !== undefined && value !== null,
		)
	);
}

/** `element` under `id` with its text taken out. */
function bodyCopy(element: PptxElement, id: string): PptxElement {
	return { ...element, id, text: '', textSegments: [] } as PptxElement;
}

/** `element` under `id` with its fill and outline taken out. */
function textCopy(element: PptxElement, id: string): PptxElement {
	const style = hasShapeProperties(element) ? (element.shapeStyle ?? {}) : {};
	const kept = Object.fromEntries(Object.entries(style).filter(([key]) => !BODY_KEYS.test(key)));
	return {
		...element,
		id,
		shapeStyle: { ...kept, fillMode: 'none', strokeWidth: 0 },
	} as PptxElement;
}

function round(value: number): number {
	return Math.round(value * 1000) / 1000;
}

function transformOf(t: MorphTextTransform, rotation: number): string {
	return `translate(${round(t.dx)}px, ${round(t.dy)}px) rotate(${round(rotation)}deg) scale(${round(t.sx)}, ${round(t.sy)})`;
}

/** The box interpolation's incoming start: the incoming box laid over the outgoing one. */
function boxStart(from: PptxElement, to: PptxElement): MorphTextTransform {
	return {
		dx: from.x + from.width / 2 - (to.x + to.width / 2),
		dy: from.y + from.height / 2 - (to.y + to.height / 2),
		sx: Math.max(from.width, 1) / Math.max(to.width, 1),
		sy: Math.max(from.height, 1) / Math.max(to.height, 1),
	};
}

/** The box interpolation's ghost end: the outgoing box laid over the incoming one. */
function boxEnd(from: PptxElement, to: PptxElement): MorphTextTransform {
	return {
		dx: to.x + to.width / 2 - (from.x + from.width / 2),
		dy: to.y + to.height / 2 - (from.y + from.height / 2),
		sx: Math.max(to.width, 1) / Math.max(from.width, 1),
		sy: Math.max(to.height, 1) / Math.max(from.height, 1),
	};
}

/** Whether two transforms place the same box within a fraction of a pixel. */
function sameTransform(a: MorphTextTransform, b: MorphTextTransform): boolean {
	return (
		Math.abs(a.dx - b.dx) < 0.5 &&
		Math.abs(a.dy - b.dy) < 0.5 &&
		Math.abs(a.sx - b.sx) < 1e-3 &&
		Math.abs(a.sy - b.sy) < 1e-3
	);
}

function splittable(pair: MorphPair, ghostIds: ReadonlySet<string>, measure?: MorphTextMeasure) {
	const { fromElement: from, toElement: to } = pair;
	if (
		!ghostIds.has(from.id) ||
		!(paintsBody(from) || paintsBody(to)) ||
		hasEffects(from) ||
		hasEffects(to) ||
		serializeShapeAdjustments(from) !== serializeShapeAdjustments(to) ||
		morphTextReplacedInSlot(from, to) ||
		morphOffstageTextSizeChange(from, to) ||
		morphOffstageTextAdoptsTarget(from, to)
	) {
		return undefined;
	}
	const alignment = resolveMorphTextAlignment(from, to, measure, { withBody: true });
	if (!alignment || sameTransform(morphTextIncomingStart(alignment, to), boxStart(from, to))) {
		return undefined;
	}
	return alignment;
}

function keyframesOf(
	name: string,
	from: { transform: string; opacity: number },
	to: { transform: string; opacity: number },
): string {
	return `
@keyframes ${name} {
\tfrom {
\t\ttransform: ${from.transform};
\t\topacity: ${from.opacity};
\t}
\tto {
\t\ttransform: ${to.transform};
\t\topacity: ${to.opacity};
\t}
}`;
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
	for (const pair of pairs) {
		const alignment = splittable(pair, context.ghostIds, measure);
		const { fromElement: from, toElement: to } = pair;
		const rising =
			alignment &&
			resolveMorphLiftCrossers(
				{ ...context, overlaid: new Set([...context.overlaid, ...plan.rising]) },
				travelledBox(from, to),
				from,
				to,
			);
		if (!alignment || !rising) {
			continue;
		}
		plan.pairs.push(pair);
		for (const id of rising) {
			plan.rising.add(id);
		}
		const fromRotation = from.rotation ?? 0;
		const toRotation = to.rotation ?? 0;
		const fromOpacity = from.opacity ?? 1;
		const toOpacity = to.opacity ?? 1;
		const still = { dx: 0, dy: 0, sx: 1, sy: 1 };
		const legs: [MorphCopyPart, MorphTextTransform, MorphTextTransform][] = [
			['body', boxEnd(from, to), boxStart(from, to)],
			['text', morphTextGhostEnd(alignment, from), morphTextIncomingStart(alignment, to)],
		];
		for (const [part, ghostEnd, incomingStart] of legs) {
			const outgoingId = morphCopyId(from.id, part);
			const incomingId = morphCopyId(to.id, part);
			const safe = (id: string) => `pptx-morph-${part}-${id.replace(/[^a-zA-Z0-9]/gu, '')}`;
			plan.keyframes.push(
				keyframesOf(
					safe(outgoingId),
					{ transform: transformOf(still, fromRotation), opacity: fromOpacity },
					{ transform: transformOf(ghostEnd, fromRotation), opacity: 0 },
				),
				keyframesOf(
					safe(incomingId),
					{ transform: transformOf(incomingStart, toRotation), opacity: 0 },
					{ transform: transformOf(still, toRotation), opacity: toOpacity },
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
			plan.sources.set(outgoingId, from.id);
			plan.sources.set(incomingId, to.id);
			plan.parts.set(outgoingId, part);
			plan.parts.set(incomingId, part);
			const copy = part === 'body' ? bodyCopy : textCopy;
			plan.groups.push({
				outgoing: copy(from, outgoingId),
				incoming: copy(to, incomingId),
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
