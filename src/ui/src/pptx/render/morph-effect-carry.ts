/**
 * A morphing shape's EFFECTS (outer shadow, glow) carried the way PowerPoint
 * carries them.
 *
 * Measured on PowerPoint 16 (`CreateVideo`, 60 fps, 720p, Morph by object,
 * 2 s) on a rectangle with a hard black outer shadow (20 pt at 45 degrees, no
 * blur) growing from 200 x 120 pt to 440 x 240 pt: mid-morph there are TWO
 * shadows, one offset `20 pt * W(p) / W0` (the outgoing shadow riding the box)
 * and one `20 pt * W(p) / W1` (the incoming one). Where the box has grown
 * 0.12, 0.23, 0.36, 0.48, 0.59, 0.68, 0.76, 0.83, 0.89, 0.93 of the way, the
 * outer one's alpha reads 0.90, 0.80, 0.67, 0.54, 0.44, 0.35, 0.27, 0.20, 0.14,
 * 0.08 and the inner one's 0.16, 0.27, 0.40, 0.53, 0.61, 0.71, 0.80, 0.85,
 * 0.92, 0.95: `1 - p` and `p` of the morph's eased progress within 0.05, and
 * where the two overlap the shadow stays fully opaque (at 50 % transparency
 * it holds 0.52; with a fill change too, the same). A glow and a soft edge
 * dissolve the same way, two renders each riding the box: a 20 pt soft edge's
 * profile across the right edge is fitted within 0.002-0.005 coverage at every
 * frame by `(1 - p) F(x / k0) + p F(x / k1)`, against 0.02-0.07 for one
 * feather of either end riding the box. A shadow whose
 * distance and blur change in place (10 -> 50 pt, 0 -> 30 pt) fades its new
 * extent in on `p`. PowerPoint cross-dissolves the two ends' effects; the body
 * is ONE shape mid-morph (see `morph-shape-carry`).
 *
 * A CSS `box-shadow` list paints both shadows on the half's own container, in
 * its local frame, plus the layers `morph-effect-shadow` adds where they
 * overlap so the stack sums like a dissolve instead of dipping. A glow is a
 * `drop-shadow()` chain fitted by `morph-effect-glow`. Both are baked as
 * keyframes on the morph's eased progress, since the overlap's alpha is not
 * linear in it.
 *
 * Under a ghost: when the outgoing ghost dissolves over a half that stays at
 * full strength (an opaque body whose fill changed), the ghost paints the
 * outgoing effects itself, fading on `1 - p`, so the incoming half only fades
 * its own in, with the overlap solved for the ghost on top. A half that fades
 * in as well is left alone: it dissolves with its ghost in one additive group
 * when the overlay lifts it (`morph-crossfade-group`), which is the measured
 * blend already.
 *
 * Not carried here: a soft edge. It is an SVG filter fixed in the element's own
 * frame, which no keyframe can reach, so a soft-edged pair dissolves its two
 * renders additively in a crossfade group instead (`morph-soft-edge-blend`).
 * Several drop-shadows on one end (a glow and a line glow)
 * have no chain to fit and run each one's radius and colour across instead.
 *
 * @module render/morph-effect-carry
 */
import type { PptxElement } from 'ooxml-core/pptx';

import { fitGlowChain, fitGlowUnderGhost } from './morph-effect-glow';
import { parseDropShadows, parseShadows, scaleLayer, withAlpha } from './morph-effect-parse';
import type { ShadowLayer, ShadowOverlap } from './morph-effect-shadow';
import {
	overlapColor,
	serializeOverlap,
	serializeShadow,
	shadowOverlap,
} from './morph-effect-shadow';
import type { MorphCarryHalf } from './morph-shape-carry';
import { bakeStops } from './morph-shape-carry-path';
import { getComputedEffectStyle } from './visual-effects';

/** Stops the effect keyframes are baked at (eased progress). */
const EFFECT_STOPS = 10;

/**
 * Whether two shapes paint different container effects (shadow, glow), so a
 * pair that does not move still has something to morph: PowerPoint fades a
 * shadow whose distance or blur changed in place (see the module notes).
 */
export function morphEffectsDiffer(from: PptxElement, to: PptxElement): boolean {
	if (from.type !== 'shape' || to.type !== 'shape') {
		return false;
	}
	const a = getComputedEffectStyle(from);
	const b = getComputedEffectStyle(to);
	return (a.boxShadow ?? '') !== (b.boxShadow ?? '') || (a.filter ?? '') !== (b.filter ?? '');
}

/** What an effect carry adds to one half's container. */
export interface MorphEffectCarry {
	/** A keyframes body (stops on the morph's eased progress), run with `linear` timing. */
	readonly keyframes: string;
}

/**
 * Who paints the outgoing effects: the half itself (`alone`), or the ghost
 * dissolving over a half that stays at full strength (`underGhost`).
 */
export type MorphEffectPainter = 'alone' | 'underGhost';

function progressStops(): number[] {
	return Array.from({ length: EFFECT_STOPS + 1 }, (_, i) => i / EFFECT_STOPS);
}

/** The `box-shadow` value at progress `p`, or `undefined` when there is nothing to carry. */
function shadowTrack(
	outA: ShadowLayer[],
	inB: ShadowLayer[],
	painter: MorphEffectPainter,
	width: number,
	height: number,
): ((p: number) => string) | undefined {
	if (painter === 'alone') {
		const still = (alpha: number, layers: ShadowLayer[]) =>
			layers.map((layer) => serializeShadow(layer, alpha)).join(', ');
		if (still(1, outA) === still(1, inB)) {
			return undefined;
		}
	} else if (inB.length === 0) {
		return undefined;
	}
	// Layer i of one end overlaps layer i of the other; lists of different
	// lengths pair nothing and fall back to the two plain fades.
	const overlaps: Array<{ overlap: ShadowOverlap; a: ShadowLayer; b: ShadowLayer }> = [];
	if (outA.length === inB.length) {
		outA.forEach((a, i) => {
			const b = inB[i];
			const overlap = b ? shadowOverlap(a, b, width, height) : undefined;
			if (b && overlap) overlaps.push({ overlap, a, b });
		});
	}
	return (p) => {
		const top = overlaps.flatMap(({ overlap, a, b }) => {
			const color = overlapColor(a.color, b.color, p, painter) ?? [0, 0, 0, 0];
			return serializeOverlap(overlap, color);
		});
		const fadeIn = inB.map((layer) => serializeShadow(layer, p));
		const fadeOut = painter === 'alone' ? outA.map((layer) => serializeShadow(layer, 1 - p)) : [];
		return [...top, ...fadeIn, ...fadeOut].join(', ');
	};
}

/** The glow `filter` at progress `p`, or `undefined` when there is nothing to carry. */
function glowTrack(
	glowA: ShadowLayer[],
	glowB: ShadowLayer[],
	painter: MorphEffectPainter,
): ((p: number) => string) | undefined {
	const radius = (layer: ShadowLayer | undefined) => layer?.lengths[2] ?? 0;
	const drop = (layer: ShadowLayer, alpha: number) =>
		`drop-shadow(${serializeShadow(withAlpha(layer, 1), alpha)})`;
	const a = glowA[0];
	const b = glowB[0];
	if (glowA.length > 1 || glowB.length > 1) {
		// Several drop-shadows (a glow and a line glow): no chain to fit, so each
		// one's radius and colour run from one end to the other.
		if (painter === 'underGhost' || glowA.length !== glowB.length) return undefined;
		const list = (layers: ShadowLayer[]) =>
			layers.map((layer) => `drop-shadow(${serializeShadow(layer, 1)})`);
		const start = list(glowA);
		const end = list(glowB);
		if (start.join(' ') === end.join(' ')) return undefined;
		return (p) =>
			glowA
				.map((layer, i) => {
					const other = glowB[i] ?? layer;
					const mix = (x: number, y: number) => x + (y - x) * p;
					return `drop-shadow(${serializeShadow(
						{
							inset: false,
							lengths: layer.lengths.map((v, k) => mix(v, other.lengths[k] ?? v)),
							color: layer.color.map((v, k) => mix(v, other.color[k] ?? v)) as ShadowLayer['color'],
						},
						1,
					)})`;
				})
				.join(' ');
	}
	if (!a && !b) {
		return undefined;
	}
	if (painter === 'underGhost') {
		if (!b) return undefined;
		const endA = { radius: radius(a), alpha: a?.color[3] ?? 0 };
		const endB = { radius: radius(b), alpha: b.color[3] };
		return (p) => drop(b, fitGlowUnderGhost(endA, endB, p));
	}
	// A glow on one end only dissolves against an invisible one of the same size.
	const layerA = a ?? withAlpha(b as ShadowLayer, 0);
	const layerB = b ?? withAlpha(layerA, 0);
	if (drop(layerA, layerA.color[3]) === drop(layerB, layerB.color[3])) {
		return undefined;
	}
	const endA = { radius: radius(layerA), alpha: layerA.color[3] };
	const endB = { radius: radius(layerB), alpha: layerB.color[3] };
	const aFirst = endA.radius <= endB.radius;
	return (p) => {
		const [narrow, wide] = fitGlowChain(endA, endB, p);
		return aFirst
			? `${drop(layerA, narrow)} ${drop(layerB, wide)}`
			: `${drop(layerB, narrow)} ${drop(layerA, wide)}`;
	};
}

/**
 * The effect carry for one half of a pair, or `undefined` when both ends'
 * effects land identically in that half's frame.
 *
 * @param painter - `alone` when this half is the only painter of the pair's
 *   effects; `underGhost` when the outgoing ghost dissolves over this
 *   (incoming, full-strength) half and paints the outgoing effects itself.
 */
export function buildMorphEffectCarry(
	from: PptxElement,
	to: PptxElement,
	half: MorphCarryHalf,
	painter: MorphEffectPainter = 'alone',
): MorphEffectCarry | undefined {
	const a = getComputedEffectStyle(from);
	const b = getComputedEffectStyle(to);
	const local = half === 'to' ? to : from;
	const width = Math.max(local.width, 1);
	const height = Math.max(local.height, 1);
	const fromK: [number, number] = [
		width / Math.max(from.width, 1),
		height / Math.max(from.height, 1),
	];
	const toK: [number, number] = [width / Math.max(to.width, 1), height / Math.max(to.height, 1)];
	const tracks: Array<(p: number) => string> = [];

	const shadowsA = parseShadows(a.boxShadow);
	const shadowsB = parseShadows(b.boxShadow);
	if (shadowsA && shadowsB && shadowsA.length + shadowsB.length > 0) {
		const track = shadowTrack(
			shadowsA.map((layer) => scaleLayer(layer, ...fromK)),
			shadowsB.map((layer) => scaleLayer(layer, ...toK)),
			painter,
			width,
			height,
		);
		if (track) tracks.push((p) => `box-shadow: ${track(p)}`);
	}

	const glowA = parseDropShadows(a.filter);
	const glowB = parseDropShadows(b.filter);
	if (glowA && glowB) {
		const track = glowTrack(
			glowA.map((layer) => scaleLayer(layer, ...fromK)),
			glowB.map((layer) => scaleLayer(layer, ...toK)),
			painter,
		);
		if (track) tracks.push((p) => `filter: ${track(p)}`);
	}
	if (tracks.length === 0) {
		return undefined;
	}
	return { keyframes: bakeStops(progressStops(), (p) => tracks.map((track) => track(p))) };
}
