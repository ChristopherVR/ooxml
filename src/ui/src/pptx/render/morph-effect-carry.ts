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
 * where the two overlap the shadow stays fully opaque. A glow (20 pt) and a blurred shadow widen mid-morph the same
 * way; a shadow whose distance and blur change in place (10 -> 50 pt, 0 -> 30
 * pt) fades its new extent in on `p`.
 *
 * The body is ONE shape mid-morph (see `morph-shape-carry`), so only the
 * effects dissolve. A CSS `box-shadow` list can paint both: the outgoing
 * layers scaled into the half's local frame fading out and the incoming ones
 * fading in, on the journey's own two frames, so the alphas run linearly in
 * eased progress like the measurement. Layers of one `box-shadow` composite
 * over each other rather than adding, so the overlap dips to 0.75 at the
 * midpoint where PowerPoint holds 1.
 *
 * A glow is also painted by a `drop-shadow()` filter; its radius is lerped
 * between the two ends' radii in the local frame (one widening glow rather
 * than two). Soft edges (an SVG filter) are not carried.
 *
 * @module render/morph-effect-carry
 */
import type { PptxElement } from 'ooxml-core/pptx';

import { splitCssList, splitCssTokens } from './css-top-level-split';
import type { MorphCarryHalf } from './morph-shape-carry';
import { getComputedEffectStyle } from './visual-effects';

interface ShadowLayer {
	inset: boolean;
	lengths: number[];
	color: [number, number, number, number];
}

const LENGTH = /^-?(?:\d+\.?\d*|\.\d+)(?:px)?$/u;

/** Parse a CSS colour this module can fade: `rgb()`, `rgba()` or hex. */
function parseColor(token: string): [number, number, number, number] | undefined {
	const fn = /^rgba?\(([^)]*)\)$/u.exec(token);
	if (fn) {
		const parts = (fn[1] ?? '')
			.split(/[\s,/]+/u)
			.filter(Boolean)
			.map(Number);
		if (parts.length < 3 || parts.some((part) => !Number.isFinite(part))) return undefined;
		return [parts[0] ?? 0, parts[1] ?? 0, parts[2] ?? 0, parts[3] ?? 1];
	}
	const hex = /^#([0-9a-f]{3,8})$/iu.exec(token)?.[1];
	if (!hex || ![3, 4, 6, 8].includes(hex.length)) return undefined;
	const full = hex.length <= 4 ? [...hex].map((c) => c + c).join('') : hex;
	const at = (i: number) => Number.parseInt(full.slice(i, i + 2), 16);
	return [at(0), at(2), at(4), full.length === 8 ? at(6) / 255 : 1];
}

function parseShadows(value: string | undefined): ShadowLayer[] | undefined {
	if (!value || value === 'none') return [];
	const layers: ShadowLayer[] = [];
	for (const layer of splitCssList(value)) {
		let inset = false;
		const lengths: number[] = [];
		let color: ShadowLayer['color'] | undefined;
		for (const token of splitCssTokens(layer.trim())) {
			if (token === 'inset') inset = true;
			else if (LENGTH.test(token)) lengths.push(Number.parseFloat(token));
			else color = parseColor(token);
		}
		if (!color || lengths.length < 2) return undefined;
		layers.push({ inset, lengths, color });
	}
	return layers;
}

function round(value: number): number {
	return Math.round(value * 1000) / 1000;
}

function serialize(layer: ShadowLayer, kx: number, ky: number, alpha: number): string {
	const mean = Math.sqrt(kx * ky);
	const lengths = layer.lengths.map(
		(value, index) => `${round(value * (index === 0 ? kx : index === 1 ? ky : mean))}px`,
	);
	const [r, g, b, a] = layer.color;
	return `${layer.inset ? 'inset ' : ''}${lengths.join(' ')} rgba(${r}, ${g}, ${b}, ${round(a * alpha)})`;
}

/** `drop-shadow()` functions of a filter, or `undefined` when it has anything else. */
function parseDropShadows(value: string | undefined): ShadowLayer[] | undefined {
	if (!value || value === 'none') return [];
	const layers: ShadowLayer[] = [];
	for (const fn of value.match(/[a-z-]+\([^()]*(?:\([^()]*\)[^()]*)*\)/gu) ?? []) {
		const inner = /^drop-shadow\((.*)\)$/u.exec(fn)?.[1];
		const parsed = inner === undefined ? undefined : parseShadows(inner);
		if (!parsed || parsed.length !== 1 || !parsed[0]) return undefined;
		layers.push(parsed[0]);
	}
	return layers;
}

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

/** What an effect carry adds to one half's container, on the journey's two frames. */
export interface MorphEffectCarry {
	readonly startDeclarations: readonly string[];
	readonly endDeclarations: readonly string[];
}

/**
 * The effect carry for one half of a pair that is the ONLY painter of its
 * pair's effects (no ghost dissolving over it), or `undefined` when both ends'
 * effects land identically in that half's frame.
 */
export function buildMorphEffectCarry(
	from: PptxElement,
	to: PptxElement,
	half: MorphCarryHalf,
): MorphEffectCarry | undefined {
	const a = getComputedEffectStyle(from);
	const b = getComputedEffectStyle(to);
	const local = half === 'to' ? to : from;
	const fromK: [number, number] = [
		Math.max(local.width, 1) / Math.max(from.width, 1),
		Math.max(local.height, 1) / Math.max(from.height, 1),
	];
	const toK: [number, number] = [
		Math.max(local.width, 1) / Math.max(to.width, 1),
		Math.max(local.height, 1) / Math.max(to.height, 1),
	];
	const startDeclarations: string[] = [];
	const endDeclarations: string[] = [];

	const shadowsA = parseShadows(a.boxShadow);
	const shadowsB = parseShadows(b.boxShadow);
	if (shadowsA && shadowsB && shadowsA.length + shadowsB.length > 0) {
		const outA = (alpha: number) => shadowsA.map((l) => serialize(l, ...fromK, alpha));
		const inB = (alpha: number) => shadowsB.map((l) => serialize(l, ...toK, alpha));
		if (outA(1).join(', ') !== inB(1).join(', ')) {
			startDeclarations.push(`box-shadow: ${[...inB(0), ...outA(1)].join(', ')}`);
			endDeclarations.push(`box-shadow: ${[...inB(1), ...outA(0)].join(', ')}`);
		}
	}

	const glowA = parseDropShadows(a.filter);
	const glowB = parseDropShadows(b.filter);
	if (glowA && glowB && glowA.length === glowB.length && glowA.length > 0) {
		const start = glowA.map((l) => `drop-shadow(${serialize(l, ...fromK, 1)})`).join(' ');
		const end = glowB.map((l) => `drop-shadow(${serialize(l, ...toK, 1)})`).join(' ');
		if (start !== end) {
			startDeclarations.push(`filter: ${start}`);
			endDeclarations.push(`filter: ${end}`);
		}
	}
	return startDeclarations.length > 0 ? { startDeclarations, endDeclarations } : undefined;
}
