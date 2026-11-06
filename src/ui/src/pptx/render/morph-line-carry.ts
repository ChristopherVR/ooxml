/**
 * A straight line or connector carried the way PowerPoint's Morph carries it.
 *
 * Measured on PowerPoint 16 (`CreateVideo`, 60 fps, 720p, Morph by object, 2 s):
 * a 4 pt line (5.3 px) moving and growing 2.2x keeps its thickness for the whole
 * morph (5.00 to 5.08 px on every frame), where our journey scales the whole
 * wrapper and so the stroke with it (1.97 px growing to 5.33 px). The line's
 * ENDS travel from one slide's endpoints to the other's, and everything along
 * the line (dash lengths, arrowheads) scales with its length, never its width.
 *
 * The journey (`translate rotate scale`) is a property of the wrapper, so it is
 * undone per stop on the painted strand: the strand gets the similarity
 * transform that lands its own line on the interpolated endpoints, expressed
 * in the wrapper's frame (`M = W(p)^-1 * Sim(p)`), and a `stroke-width` that
 * the similarity's scale brings back to the line's width. Arrowhead markers
 * (sized in stroke widths) and dashes (user units) scale with the line.
 *
 * Both halves of a pair get the same line, so a ghost and its incoming half
 * stay on top of each other.
 *
 * @module render/morph-line-carry
 */
import type { PptxElement } from 'ooxml-core/pptx';
import { hasShapeProperties } from 'ooxml-core/pptx';

import { connectorNeedsPath, getCompoundLineOffsets } from './connector-style';
import { reachOffsets } from './morph-line-reach';
import { shortestRotationTarget } from './morph-rotation';
import { bakeStops, lerp, round } from './morph-shape-carry-path';
import type { MorphAnimationStyle, MorphPair } from './morph-types';

/** Which half of a pair a carry is painted on: its box is the local frame. */
export type MorphLineHalf = 'from' | 'to';

/** The baked keyframes body for the painted strand of a line. */
export interface MorphLineCarry {
	readonly keyframes: string;
}

/** `[a, b, c, d, e, f]`: CSS `matrix()` order (`x' = a x + c y + e`, `y' = b x + d y + f`). */
export type Matrix = readonly [number, number, number, number, number, number];
export type Vec = readonly [number, number];

/** A matrix entry to five places: a rotation's sine is a long lever over a long line. */
export const precise = (value: number): number => Math.round(value * 1e5) / 1e5;

export const IDENTITY: Matrix = [1, 0, 0, 1, 0, 0];
export const STOPS = 24;
export const EPSILON = 0.5;
export const DEG = Math.PI / 180;

export function mul(m: Matrix, n: Matrix): Matrix {
	return [
		m[0] * n[0] + m[2] * n[1],
		m[1] * n[0] + m[3] * n[1],
		m[0] * n[2] + m[2] * n[3],
		m[1] * n[2] + m[3] * n[3],
		m[0] * n[4] + m[2] * n[5] + m[4],
		m[1] * n[4] + m[3] * n[5] + m[5],
	];
}

export function invert(m: Matrix): Matrix | undefined {
	const det = m[0] * m[3] - m[1] * m[2];
	if (Math.abs(det) < 1e-9) return undefined;
	const a = m[3] / det;
	const b = -m[1] / det;
	const c = -m[2] / det;
	const d = m[0] / det;
	return [a, b, c, d, -(a * m[4] + c * m[5]), -(b * m[4] + d * m[5])];
}

export const translate = (x: number, y: number): Matrix => [1, 0, 0, 1, x, y];
export const scale = (x: number, y: number): Matrix => [x, 0, 0, y, 0, 0];
export function rotate(deg: number): Matrix {
	const s = Math.sin(deg * DEG);
	const c = Math.cos(deg * DEG);
	return [c, s, -s, c, 0, 0];
}
export function apply(m: Matrix, v: Vec): Vec {
	return [m[0] * v[0] + m[2] * v[1] + m[4], m[1] * v[0] + m[3] * v[1] + m[5]];
}

/** The journey of one half's wrapper: its own box, and the keyframe values at each end. */
export interface Journey {
	readonly local: PptxElement;
	readonly shift: readonly [Vec, Vec];
	readonly turn: readonly [number, number];
	readonly stretch: readonly [Vec, Vec];
}

export function centre(el: PptxElement): Vec {
	return [el.x + el.width / 2, el.y + el.height / 2];
}

export function journeyOf(from: PptxElement, to: PptxElement, half: MorphLineHalf): Journey {
	const [fcx, fcy] = centre(from);
	const [tcx, tcy] = centre(to);
	const fromRot = from.rotation ?? 0;
	const toRot = to.rotation ?? 0;
	const sx = Math.max(to.width, 1) / Math.max(from.width, 1);
	const sy = Math.max(to.height, 1) / Math.max(from.height, 1);
	// The two halves of `morph-animation`, as the plan emits them.
	return half === 'to'
		? {
				local: to,
				shift: [
					[fcx - tcx, fcy - tcy],
					[0, 0],
				],
				turn: [shortestRotationTarget(toRot, fromRot), toRot],
				stretch: [
					[1 / sx, 1 / sy],
					[1, 1],
				],
			}
		: {
				local: from,
				shift: [
					[0, 0],
					[tcx - fcx, tcy - fcy],
				],
				turn: [fromRot, shortestRotationTarget(fromRot, toRot)],
				stretch: [
					[1, 1],
					[sx, sy],
				],
			};
}

/** The wrapper's local-to-slide matrix at eased progress `p`. */
export function wrapperAt(journey: Journey, p: number): Matrix {
	const { local } = journey;
	const [cx, cy] = [local.width / 2, local.height / 2];
	const [tx, ty] = [
		lerp(journey.shift[0][0], journey.shift[1][0], p),
		lerp(journey.shift[0][1], journey.shift[1][1], p),
	];
	const [sx, sy] = [
		lerp(journey.stretch[0][0], journey.stretch[1][0], p),
		lerp(journey.stretch[0][1], journey.stretch[1][1], p),
	];
	return mul(
		translate(local.x + cx, local.y + cy),
		mul(
			translate(tx, ty),
			mul(
				rotate(lerp(journey.turn[0], journey.turn[1], p)),
				mul(scale(sx, sy), translate(-cx, -cy)),
			),
		),
	);
}

/** A line's two end points in the strand's own frame (flips baked in, as every binding draws it). */
function strandEnds(el: PptxElement): [Vec, Vec] {
	const w = Math.max(el.width, 1);
	const h = Math.max(el.height, 1);
	return [
		[el.flipHorizontal ? w : 0, el.flipVertical ? h : 0],
		[el.flipHorizontal ? 0 : w, el.flipVertical ? 0 : h],
	];
}

/** A line's two end points on the slide: its box (rotated about its centre) and its flips. */
function slideEnds(el: PptxElement): [Vec, Vec] {
	const [cx, cy] = centre(el);
	const to = mul(
		translate(cx, cy),
		mul(rotate(el.rotation ?? 0), translate(-el.width / 2, -el.height / 2)),
	);
	const w = el.width;
	const h = el.height;
	return [
		apply(to, [el.flipHorizontal ? w : 0, el.flipVertical ? h : 0]),
		apply(to, [el.flipHorizontal ? 0 : w, el.flipVertical ? 0 : h]),
	];
}

export function lineWidth(el: PptxElement): number {
	return hasShapeProperties(el) ? Math.max(0, el.shapeStyle?.strokeWidth ?? 2) : 2;
}

/** Whether `element` is a straight, single-strand connector or line the carry can re-path. */
export function isCarriedLine(element: PptxElement): boolean {
	if (element.type !== 'connector' || !hasShapeProperties(element)) return false;
	const style = element.shapeStyle;
	return (
		!connectorNeedsPath(element.shapeType) &&
		lineWidth(element) > 0 &&
		getCompoundLineOffsets(style?.compoundLine, lineWidth(element)).length === 1
	);
}

/**
 * The carry for one half of a pair of straight lines whose box, rotation or
 * flips differ, or `undefined` when the pair is not one.
 */
export function buildMorphLineCarry(
	from: PptxElement,
	to: PptxElement,
	half: MorphLineHalf,
): MorphLineCarry | undefined {
	if (!isCarriedLine(from) || !isCarriedLine(to)) return undefined;
	const unchanged =
		Math.abs(from.width - to.width) <= EPSILON &&
		Math.abs(from.height - to.height) <= EPSILON &&
		(from.rotation ?? 0) === (to.rotation ?? 0) &&
		Boolean(from.flipHorizontal) === Boolean(to.flipHorizontal) &&
		Boolean(from.flipVertical) === Boolean(to.flipVertical);
	const journey = journeyOf(from, to, half);
	const [a, b] = strandEnds(journey.local);
	const localLength = Math.hypot(b[0] - a[0], b[1] - a[1]);
	const [fa, fb] = slideEnds(from);
	const [ta, tb] = slideEnds(to);
	if (unchanged || localLength < 1e-6) return undefined;
	const w0 = lineWidth(from);
	const w1 = lineWidth(to);
	const angle = Math.atan2(b[1] - a[1], b[0] - a[0]);
	const stops: Array<{ p: number; matrix: Matrix; width: number }> = [];
	for (let i = 0; i <= STOPS; i++) {
		const p = i / STOPS;
		const linearStart: Vec = [lerp(fa[0], ta[0], p), lerp(fa[1], ta[1], p)];
		const linearEnd: Vec = [lerp(fb[0], tb[0], p), lerp(fb[1], tb[1], p)];
		// PowerPoint grows the box about its (linearly moving) centre ahead of `p`.
		const [reachX, reachY] = reachOffsets(from, to, p);
		const mid: Vec = [(linearStart[0] + linearEnd[0]) / 2, (linearStart[1] + linearEnd[1]) / 2];
		const half: Vec = [
			(linearEnd[0] - linearStart[0] + reachX) / 2,
			(linearEnd[1] - linearStart[1] + reachY) / 2,
		];
		const start: Vec = [mid[0] - half[0], mid[1] - half[1]];
		const end: Vec = [mid[0] + half[0], mid[1] + half[1]];
		const length = Math.hypot(end[0] - start[0], end[1] - start[1]);
		const wrapperInverse = invert(wrapperAt(journey, p));
		if (length < 1 || !wrapperInverse) return undefined;
		const k = length / localLength;
		const turn = Math.atan2(end[1] - start[1], end[0] - start[0]) - angle;
		// Sim(q) = start + k R(turn) (q - a), as a matrix.
		const sim = mul(
			translate(start[0], start[1]),
			mul(rotate(turn / DEG), mul(scale(k, k), translate(-a[0], -a[1]))),
		);
		stops.push({ p, matrix: mul(wrapperInverse, sim), width: lerp(w0, w1, p) / k });
	}
	const progress = stops.map((stop) => stop.p);
	const keyframes = bakeStops(progress, (p) => {
		const stop = stops[progress.indexOf(p)] ?? {
			matrix: IDENTITY,
			width: lineWidth(journey.local),
		};
		return [
			'transform-origin: 0px 0px',
			'vector-effect: none',
			`transform: matrix(${stop.matrix.map(precise).join(', ')})`,
			`stroke-width: ${round(stop.width)}px`,
		];
	});
	return { keyframes };
}

/** The animation of one half's painted strand (`target: 'line'`), or none. */
export function buildMorphLineAnimations(
	pair: MorphPair,
	id: string,
	half: MorphLineHalf,
	durationMs: number,
	tag: string,
): MorphAnimationStyle[] {
	const carry = buildMorphLineCarry(pair.fromElement, pair.toElement, half);
	if (!carry) return [];
	const name = `pptx-morph-line-${tag}-${id.replace(/[^a-zA-Z0-9]/gu, '')}`;
	return [
		{
			elementId: id,
			animation: `${name} ${durationMs}ms linear forwards`,
			keyframes: `\n@keyframes ${name} {\n${carry.keyframes}\n}`,
			target: 'line',
		},
	];
}
