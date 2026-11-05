/**
 * The glyph copies that carry a text box turning while its type changes.
 *
 * Measured on PowerPoint 16 (`CreateVideo`, 60 fps, 720p) on a 32 pt Calibri
 * line turning 0 -> 60 degrees while it becomes 40 pt Arial (and on two more
 * pairs: the same line turning 0 -> 40 degrees into 32 pt Arial, and 28 pt
 * Calibri turning 0 -> 90 degrees into 44 pt): PowerPoint does not turn the
 * line as one piece. Every glyph keeps its own shape, grows on the morph's
 * progress `p`, turns about ITS OWN middle by the box's angle, and that
 * middle travels in a straight line from where the glyph stands on the first
 * slide to where it stands on the second. The line therefore follows the
 * CHORD between its two layouts, not the arc: its length read 331, 375 and
 * 401 px at 49%, 79% and 90% of the turn (a chord of the two layouts gives
 * 326, 379 and 406; a rigid turn would hold 318 to 432 on a straight lerp),
 * it dips below its start length while it turns (Calibri -> Arial 32 pt, 0 ->
 * 40 degrees: 318 to 314 px), its words run into each other mid-morph while
 * the glyphs are already larger, and a 28 -> 44 pt line turning a quarter
 * bows. One transform for the whole line cannot do any of that.
 *
 * Each glyph is a copy of BOTH halves clipped to the glyph's advance cell. CSS
 * cannot compose a straight path with a turn on one timing function, so every
 * copy states frames sampled on the morph's own curve and plays them linearly.
 *
 * @module render/morph-text-glyph-plan
 */
import type { PptxElement } from 'ooxml-core/pptx';

import { shortestRotationTarget } from './morph-animation';
import { easedProgress } from './morph-easing';
import type { MorphTextAlignment, MorphTextMeasure } from './morph-text-frame';
import { measureMorphTextWithCanvas, resolveMorphTextAlignment } from './morph-text-frame';
import { paintsBody } from './morph-text-frame-read';
import { morphTextGlyphCells } from './morph-text-glyphs';
import type { MorphGlyphCell } from './morph-text-glyphs';
import type { MorphTextInkMeasure } from './morph-text-ink';
import { measureMorphTextInkWithCanvas } from './morph-text-ink';
import type { MorphTextPiecePlan } from './morph-text-piece-plan';
import {
	morphOffstageTextAdoptsTarget,
	morphOffstageTextSizeChange,
	morphTextReplacedInSlot,
} from './morph-text-slot';
import type { MorphPoint } from './morph-text-transform';
import type { MorphPair } from './morph-types';
import { MORPH_EASING_POINTS } from './morph-types';

/** Frames sampled on the morph's curve. */
const SAMPLES = 24;
/** A pair scaled by less than this on both axes needs no glyph copies. */
const NEUTRAL_SCALE = 0.01;
/** The clip's reach (px) at the line's ends and above and below it. */
const REACH_PX = 2000;

function round(value: number): number {
	return Math.round(value * 1000) / 1000;
}

type Turn = NonNullable<MorphTextAlignment['turn']>;
type Centres = NonNullable<MorphTextAlignment['centres']>;

/** A turning pair, resolved for glyph copies. */
export interface MorphGlyphTurn {
	alignment: MorphTextAlignment & { turn: Turn; centres: Centres };
	cellsFrom: MorphGlyphCell[];
	cellsTo: MorphGlyphCell[];
}

/** Resolve `pair` for glyph copies, or `undefined` when its whole-box turn stands. */
export function resolveMorphGlyphTurn(
	pair: MorphPair,
	ghostIds: ReadonlySet<string>,
	measure: MorphTextMeasure = measureMorphTextWithCanvas,
	ink: MorphTextInkMeasure | undefined = measure === measureMorphTextWithCanvas
		? measureMorphTextInkWithCanvas
		: undefined,
): MorphGlyphTurn | undefined {
	const { fromElement: from, toElement: to } = pair;
	if (
		!ghostIds.has(from.id) ||
		from.flipHorizontal ||
		from.flipVertical ||
		to.flipHorizontal ||
		to.flipVertical ||
		paintsBody(from) ||
		paintsBody(to) ||
		morphTextReplacedInSlot(from, to) ||
		morphOffstageTextSizeChange(from, to) ||
		morphOffstageTextAdoptsTarget(from, to)
	) {
		return undefined;
	}
	const alignment = resolveMorphTextAlignment(from, to, measure);
	if (
		!alignment?.turn ||
		!alignment.centres ||
		(Math.abs(alignment.scaleX - 1) < NEUTRAL_SCALE &&
			Math.abs(alignment.scaleY - 1) < NEUTRAL_SCALE)
	) {
		return undefined;
	}
	const cellsFrom = morphTextGlyphCells(from, measure, ink);
	const cellsTo = morphTextGlyphCells(to, measure, ink);
	if (
		!cellsFrom ||
		!cellsTo ||
		cellsFrom.length !== cellsTo.length ||
		cellsFrom.some((cell, index) => cell.index !== cellsTo[index]?.index)
	) {
		return undefined;
	}
	return {
		alignment: alignment as MorphGlyphTurn['alignment'],
		cellsFrom,
		cellsTo,
	};
}

/** Where a glyph stands, and at what angle, at progress `progress`. */
function glyphAt(
	turn: MorphGlyphTurn,
	anchors: { from: MorphPoint; to: MorphPoint },
	progress: number,
) {
	const { alignment } = turn;
	const target = shortestRotationTarget(alignment.turn.from, alignment.turn.to);
	const turned = (point: MorphPoint, centre: MorphPoint, degrees: number) => {
		const radians = (degrees * Math.PI) / 180;
		const dx = point.x - centre.x;
		const dy = point.y - centre.y;
		return {
			x: centre.x + dx * Math.cos(radians) - dy * Math.sin(radians),
			y: centre.y + dx * Math.sin(radians) + dy * Math.cos(radians),
		};
	};
	const a = turned(anchors.from, alignment.centres.from, alignment.turn.from);
	const b = turned(anchors.to, alignment.centres.to, target);
	return {
		x: a.x + (b.x - a.x) * progress,
		y: a.y + (b.y - a.y) * progress,
		rotation: alignment.turn.from + (target - alignment.turn.from) * progress,
		size: 1 + (alignment.scaleY - 1) * progress,
	};
}

function transformOf(
	at: { x: number; y: number; rotation: number },
	scale: number,
	element: PptxElement,
	anchor: MorphPoint,
): string {
	const cx = element.x + element.width / 2;
	const cy = element.y + element.height / 2;
	return `translate(${round(at.x - cx)}px, ${round(at.y - cy)}px) rotate(${round(at.rotation)}deg) scale(${round(scale)}) translate(${round(cx - anchor.x)}px, ${round(cy - anchor.y)}px)`;
}

function clipOf(cell: MorphGlyphCell, first: boolean, last: boolean): string {
	const left = first ? -REACH_PX : cell.x0;
	const right = last ? REACH_PX : cell.x1;
	return `polygon(${round(left)}px -${REACH_PX}px, ${round(right)}px -${REACH_PX}px, ${round(right)}px ${REACH_PX}px, ${round(left)}px ${REACH_PX}px)`;
}

/** The id a glyph copy is painted under. */
export function morphGlyphId(sourceId: string, index: number): string {
	return `${sourceId}-morph-glyph-${index}`;
}

function percent(value: number): string {
	return `${Math.round(value * 10000) / 100}%`;
}

/**
 * The glyph copies of every turning pair whose type changes with its angle.
 *
 * @param pairs - The matched pairs (object mode only).
 * @param ghostIds - Outgoing ids the overlay paints.
 * @param durationMs - The morph's duration.
 * @param measure - Text measurement; defaults to a canvas.
 */
export function buildMorphTextGlyphPlan(
	pairs: readonly MorphPair[],
	ghostIds: ReadonlySet<string>,
	durationMs: number,
	measure?: MorphTextMeasure,
): MorphTextPiecePlan {
	const plan: MorphTextPiecePlan = {
		pairs: [],
		groups: [],
		outgoingAnimations: new Map(),
		incomingAnimations: new Map(),
		keyframes: [],
		sources: new Map(),
	};
	const [x1, y1, x2, y2] = MORPH_EASING_POINTS;
	const progresses = Array.from({ length: SAMPLES + 1 }, (_, i) =>
		easedProgress(i / SAMPLES, x1, y1, x2, y2),
	);
	for (const pair of pairs) {
		const turn = resolveMorphGlyphTurn(pair, ghostIds, measure);
		if (!turn) {
			continue;
		}
		const { fromElement: from, toElement: to } = pair;
		plan.pairs.push(pair);
		const fromOpacity = from.opacity ?? 1;
		const toOpacity = to.opacity ?? 1;
		turn.cellsFrom.forEach((cellFrom, index) => {
			const cellTo = turn.cellsTo[index] as MorphGlyphCell;
			const first = index === 0;
			const last = index === turn.cellsFrom.length - 1;
			const outgoingId = morphGlyphId(from.id, index);
			const incomingId = morphGlyphId(to.id, index);
			const safe = (id: string) => `pptx-morph-glyph-${id.replace(/[^a-zA-Z0-9]/gu, '')}`;
			const anchorFrom = { x: from.x + cellFrom.cx, y: from.y + cellFrom.cy };
			const anchorTo = { x: to.x + cellTo.cx, y: to.y + cellTo.cy };
			const stops = progresses.map((progress) =>
				glyphAt(turn, { from: anchorFrom, to: anchorTo }, progress),
			);
			const frames = (
				name: string,
				clip: string,
				transform: (at: (typeof stops)[number]) => string,
				opacity: (progress: number) => number,
			) =>
				`\n@keyframes ${name} {\n${stops
					.map(
						(at, i) =>
							`\t${percent(i / SAMPLES)} {\n\t\ttransform: ${transform(at)};\n\t\tclip-path: ${clip};\n\t\topacity: ${round(opacity(progresses[i] as number))};\n\t}`,
					)
					.join('\n')}\n}`;
			const scaleTo = turn.alignment.scaleY;
			plan.keyframes.push(
				frames(
					safe(outgoingId),
					clipOf(cellFrom, first, last),
					(at) => transformOf(at, at.size, from, anchorFrom),
					(progress) => fromOpacity * (1 - progress),
				),
				frames(
					safe(incomingId),
					clipOf(cellTo, first, last),
					(at) => transformOf(at, at.size / scaleTo, to, anchorTo),
					(progress) => toOpacity * progress,
				),
			);
			plan.outgoingAnimations.set(
				outgoingId,
				`${safe(outgoingId)} ${durationMs}ms linear forwards`,
			);
			plan.incomingAnimations.set(
				incomingId,
				`${safe(incomingId)} ${durationMs}ms linear forwards`,
			);
			plan.sources.set(outgoingId, from.id);
			plan.sources.set(incomingId, to.id);
			plan.groups.push({
				outgoing: { ...from, id: outgoingId } as PptxElement,
				incoming: { ...to, id: incomingId } as PptxElement,
			});
		});
	}
	return plan;
}
