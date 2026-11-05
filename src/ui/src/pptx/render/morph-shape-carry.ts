/**
 * A resizing shape's GEOMETRY carried the way PowerPoint's Morph carries it.
 *
 * Every morphing element travels on a FLIP journey: it is painted at one end's
 * box and scaled onto the other's. That is exact for a picture, but not for a
 * vector shape, whose outline, corner radii and adjustment handles are a
 * function of its box. Measured on PowerPoint 16 (`CreateVideo`, 60 fps, 720p,
 * Morph by object, 2 s), PowerPoint interpolates the shape's rendered geometry
 * in BOX-NORMALISED coordinates and maps it onto the interpolated box:
 *
 *  - a rounded rectangle with a 6 pt (8 px) outline growing from 60 x 60 pt to
 *    480 x 240 pt keeps no fixed stroke width: its vertical edges read 8, 9.6,
 *    13.1, 16.1, 18.0, 18.6, 18.0, 16.8, 14.9, 12.5, 10.2, 8.4, 7.8 px
 *    across the morph and its horizontal ones 8 to 12.0 and back;
 *    `W(p) * lerp(w / W0, w / W1, p)` predicts every one of those within
 *    0.3 to 0.6 px (`H(p) * lerp(w / H0, w / H1, p)` within 0.4 px). One
 *    solid stroke, never two crossfading;
 *  - the same shape growing only 1.4x (300 x 120 to 420 x 180 pt) therefore
 *    keeps its outline within 0.3 px of 8 px on every frame;
 *  - an adjustment change (a rounded rectangle's corner 0.05 -> 0.5, an
 *    arrow's shaft and head 0.3 -> 0.7 / 0.9) morphs as one outline whose
 *    vertices move linearly, never a dissolve.
 *
 * Our journey scaled the INCOMING shape, so its geometry was the incoming
 * one's for the whole flight: the outline was `w * W(p) / W1` (1.1 px instead
 * of 8 at the start of the grow) and the corners the final ones, stretched.
 *
 * This module bakes the measured model onto the element's own nodes, in the
 * element's LOCAL frame (the box the journey scales):
 *
 *  - a rounded rectangle's `border-radius` lerps between the two normalised
 *    radii, on the journey's own two frames (linear in eased progress);
 *  - a clipped preset's `clip-path` is re-evaluated per stop;
 *  - the outline path (`[data-pptx-outline] path`) gets a `d` per stop and a
 *    `transform: scale(ax, ay)` that undoes the journey's anisotropic scale on
 *    its stroke, so the stroke is `lerp(w0 / W0, w1 / W1, p)` of the box on
 *    each axis, as measured.
 *
 * Both halves of a pair get the same geometry, so a dissolving pair's ghost
 * and incoming half stay on top of each other.
 *
 * @module render/morph-shape-carry
 */
import type { PptxElement } from 'ooxml-core/pptx';
import { getRoundRectRadiusPx, getShapeType, hasShapeProperties } from 'ooxml-core/pptx';

import {
	bakeStops,
	geometryAt,
	lerp,
	outlineCommands,
	round,
	sameStructure,
	scalePath,
	stopProgress,
} from './morph-shape-carry-path';
import { resolveShapeGeometry } from './shape-geometry-cascade';
import { outsetClipPathForStroke } from './stroke-clip-outset';
import { buildStrokeOutline, outlineStrokeJoin, straddlingStrokeWidth } from './stroke-outline';

/** Which half of a pair a carry is painted on: its box is the local frame. */
export type MorphCarryHalf = 'from' | 'to';

/** What a carry adds to one half. */
export interface MorphShapeCarry {
	/** Declarations for the half's container at the start (`from`) and end (`to`) of the morph. */
	readonly startDeclarations: readonly string[];
	readonly endDeclarations: readonly string[];
	/** Baked `clip-path` keyframes body for the container, when the preset is clipped. */
	readonly clipKeyframes?: string;
	/** Baked keyframes body for the outline path, when the outline is carried. */
	readonly outlineKeyframes?: string;
}

const EPSILON = 0.5;

/** Whether a pair is one preset shape whose box or handles changed. */
export function morphShapeCarryApplies(from: PptxElement, to: PptxElement): boolean {
	if (
		from.type !== 'shape' ||
		to.type !== 'shape' ||
		!hasShapeProperties(from) ||
		!hasShapeProperties(to) ||
		from.pathData ||
		to.pathData ||
		Boolean(from.flipHorizontal) !== Boolean(to.flipHorizontal) ||
		Boolean(from.flipVertical) !== Boolean(to.flipVertical)
	) {
		return false;
	}
	const type = getShapeType(from.shapeType);
	if (!type || type !== getShapeType(to.shapeType)) {
		return false;
	}
	const kind = resolveShapeGeometry(from).kind;
	if (kind === 'bare' || kind === 'strokeOnly' || kind === 'lineEdge') {
		return false;
	}
	const resized =
		Math.abs(from.width - to.width) > EPSILON || Math.abs(from.height - to.height) > EPSILON;
	const handles =
		JSON.stringify(from.shapeAdjustments ?? {}) !== JSON.stringify(to.shapeAdjustments ?? {});
	if (!resized && !handles) {
		return false;
	}
	const a = outlineCommands(from);
	const b = outlineCommands(to);
	return Boolean(a && b && sameStructure(a, b));
}

/** The authored outline width of a shape, 0 when it has none. */
function lineWidth(element: PptxElement): number {
	return hasShapeProperties(element) ? Math.max(0, element.shapeStyle?.strokeWidth ?? 0) : 0;
}

/**
 * Whether `element`'s outline overlay is one plain line the carry can re-path.
 * A clipped preset adds a hairline on its outer edge (`clipEdgeBleed`); the
 * re-path draws it on the line itself, where the line hides it.
 */
function carriesOutline(element: PptxElement): boolean {
	const outline = buildStrokeOutline(element);
	return Boolean(
		outline &&
		!outline.markers &&
		outline.strands[0]?.d === undefined &&
		outline.strands.every((strand) => strand.offset === 0),
	);
}

/**
 * The carry for one half of a pair, in that half's local frame, or `undefined`
 * when the pair is not a resizing or re-handled preset shape.
 */
export function buildMorphShapeCarry(
	from: PptxElement,
	to: PptxElement,
	half: MorphCarryHalf,
): MorphShapeCarry | undefined {
	if (!morphShapeCarryApplies(from, to) || !hasShapeProperties(from) || !hasShapeProperties(to)) {
		return undefined;
	}
	const a = outlineCommands(from);
	const b = outlineCommands(to);
	if (!a || !b) {
		return undefined;
	}
	const local = half === 'to' ? to : from;
	const W = Math.max(local.width, 1);
	const H = Math.max(local.height, 1);
	const startDeclarations: string[] = [];
	const endDeclarations: string[] = [];
	if (getShapeType(to.shapeType) === 'roundRect') {
		const r0 = getRoundRectRadiusPx(from);
		const r1 = getRoundRectRadiusPx(to);
		const radius = (p: number) =>
			`border-radius: ${round(lerp(r0 / from.width, r1 / to.width, p) * W)}px / ${round(
				lerp(r0 / from.height, r1 / to.height, p) * H,
			)}px`;
		startDeclarations.push(radius(0));
		endDeclarations.push(radius(1));
	}
	// The stroke, as a fraction of the box on each axis, lerps between the ends.
	const w0 = lineWidth(from);
	const w1 = lineWidth(to);
	const wl = lineWidth(local);
	const stroked = w0 > 0 && w1 > 0 && carriesOutline(local);
	// The path's own scale: the journey maps local px onto the box, this maps
	// the path's user units onto local px so its stroke lands at the measured
	// width. `1` leaves the stroke to the journey, as before.
	const scaleAt = (p: number): [number, number] =>
		stroked
			? [
					(lerp(w0 / from.width, w1 / to.width, p) * W) / wl,
					(lerp(w0 / from.height, w1 / to.height, p) * H) / wl,
				]
			: [1, 1];
	const decision = resolveShapeGeometry(local);
	const straddle = straddlingStrokeWidth(local);
	// The axis whose outline share changes most decides where the stops crowd.
	const spread = (u0: number, u1: number) => Math.max(u0 / u1, u1 / u0);
	const progress = !stroked
		? stopProgress(1, 1)
		: spread(w0 / from.width, w1 / to.width) >= spread(w0 / from.height, w1 / to.height)
			? stopProgress(w0 / from.width, w1 / to.width)
			: stopProgress(w0 / from.height, w1 / to.height);
	const clipKeyframes =
		decision.kind === 'clipPath'
			? bakeStops(progress, (p) => {
					const [ax, ay] = scaleAt(p);
					const user = geometryAt(a, b, from, to, p, W / ax, H / ay);
					const widened =
						straddle > 0
							? outsetClipPathForStroke(
									`path('${user}')`,
									straddle,
									W / ax,
									H / ay,
									outlineStrokeJoin(local),
								)
							: `path('${user}')`;
					const inner = /^path\('(.*)'\)$/su.exec(widened)?.[1] ?? user;
					return [`clip-path: path('${scalePath(inner, ax, ay)}')`];
				})
			: undefined;
	const outlineKeyframes = stroked
		? bakeStops(progress, (p) => {
				const [ax, ay] = scaleAt(p);
				return [
					`d: path('${geometryAt(a, b, from, to, p, W / ax, H / ay)}')`,
					'transform-origin: 0px 0px',
					`transform: scale(${round(ax)}, ${round(ay)})`,
				];
			})
		: undefined;
	if (!startDeclarations.length && !clipKeyframes && !outlineKeyframes) {
		return undefined;
	}
	return {
		startDeclarations,
		endDeclarations,
		...(clipKeyframes ? { clipKeyframes } : {}),
		...(outlineKeyframes ? { outlineKeyframes } : {}),
	};
}
