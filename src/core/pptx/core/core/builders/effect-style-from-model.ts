/**
 * ShapeStyle effect fields from the shared typed effects (`drawingml/effect-read.ts`).
 *
 * The pptx effect codecs read each effect through the shared reader over the object-tree view
 * and only map the typed fields onto ShapeStyle here (pixels from EMU, degrees, the validated
 * alignment). Colours and opacity still resolve through the runtime colour codec on the effect's
 * node, so scheme colours follow the slide's theme and colour map. `a:effectLst` and the effect
 * DAG (`PptxEffectDagExtractor`) both map through these functions.
 */
import {
	readDrawingEffect,
	type DrawingEffectFields,
	type DrawingEffectKind,
} from '../../../../drawingml/effect-read';
import { OBJECT_DRAWING_VIEW } from '../../color/object-drawing-view';
import type { ShapeStyle, XmlObject } from '../../types';

type Effect<K extends DrawingEffectKind> = Extract<DrawingEffectFields, { kind: K }>;

export interface EffectStyleContext {
	emuPerPx: number;
	parseColor: (colorNode: XmlObject | undefined, placeholderColor?: string) => string | undefined;
	extractColorOpacity: (colorNode: XmlObject | undefined) => number | undefined;
}

const LOCAL_NAMES: Record<DrawingEffectKind, string> = {
	outerShadow: 'outerShdw',
	innerShadow: 'innerShdw',
	presetShadow: 'prstShdw',
	glow: 'glow',
	softEdge: 'softEdge',
	blur: 'blur',
	reflection: 'reflection',
	fillOverlay: 'fillOverlay',
};

const VALID_ALIGNMENTS = new Set(['tl', 't', 'tr', 'l', 'ctr', 'r', 'bl', 'b', 'br']);

/** The first effect of `kind` in a list or container node, read through the shared reader. */
export function findObjectEffect<K extends DrawingEffectKind>(
	list: XmlObject | undefined,
	kind: K,
): { effect: Effect<K>; node: XmlObject } | undefined {
	if (!list) {
		return undefined;
	}
	const name = LOCAL_NAMES[kind];
	const node = OBJECT_DRAWING_VIEW.child(list, name);
	if (!node) {
		return undefined;
	}
	const effect = readDrawingEffect(OBJECT_DRAWING_VIEW, name, node) as Effect<K> | undefined;
	return effect ? { effect, node } : undefined;
}

/** The typed effect of an already located node (a DAG leaf). */
export function readObjectEffect<K extends DrawingEffectKind>(kind: K, node: XmlObject): Effect<K> {
	return readDrawingEffect(OBJECT_DRAWING_VIEW, LOCAL_NAMES[kind], node) as Effect<K>;
}

const nonNegativePx = (emu: number | undefined, emuPerPx: number) =>
	emu !== undefined && emu >= 0 ? emu / emuPerPx : undefined;
const integer = (value: number | undefined) =>
	value === undefined ? undefined : Math.trunc(value);

/** The x/y offset of a shadow `distance` px away at `degrees` (clockwise from +x), to 0.01 px. */
export function shadowOffset(
	distance: number | undefined,
	degrees: number,
): { x: number | undefined; y: number | undefined } {
	if (distance === undefined) {
		return { x: undefined, y: undefined };
	}
	const radians = (degrees * Math.PI) / 180;
	return {
		x: Math.round(Math.cos(radians) * distance * 100) / 100,
		y: Math.round(Math.sin(radians) * distance * 100) / 100,
	};
}

function alignment<T>(value: string | undefined): T | undefined {
	const token = String(value ?? '').trim();
	return VALID_ALIGNMENTS.has(token) ? (token as T) : undefined;
}

/** `a:outerShdw` (20.1.8.45). */
export function outerShadowStyle(
	effect: Effect<'outerShadow'>,
	node: XmlObject,
	context: EffectStyleContext,
): Partial<ShapeStyle> {
	const distance = nonNegativePx(integer(effect.distanceEmu), context.emuPerPx);
	const directionDegrees =
		effect.directionAngle !== undefined ? integer(effect.directionAngle)! / 60000 : 0;
	const offset = shadowOffset(distance, directionDegrees);
	return {
		shadowColor: context.parseColor(node),
		shadowOpacity: context.extractColorOpacity(node),
		shadowBlur: nonNegativePx(integer(effect.blurRadiusEmu), context.emuPerPx),
		shadowOffsetX: offset.x,
		shadowOffsetY: offset.y,
		shadowAngle: directionDegrees,
		shadowDistance: distance,
		shadowRotateWithShape: effect.rotateWithShape,
		shadowScaleX: integer(effect.scaleX),
		shadowScaleY: integer(effect.scaleY),
		shadowSkewX: integer(effect.skewXAngle),
		shadowSkewY: integer(effect.skewYAngle),
		shadowAlignment: alignment<ShapeStyle['shadowAlignment']>(effect.alignment),
	};
}

/** `a:innerShdw` (20.1.8.40). `rotWithShape` is outside the schema but kept when written. */
export function innerShadowStyle(
	effect: Effect<'innerShadow'>,
	node: XmlObject,
	context: EffectStyleContext,
): Partial<ShapeStyle> {
	const distance = nonNegativePx(integer(effect.distanceEmu), context.emuPerPx);
	const directionDegrees =
		effect.directionAngle !== undefined ? integer(effect.directionAngle)! / 60000 : 0;
	const offset = shadowOffset(distance, directionDegrees);
	const rotateWithShape = node['@_rotWithShape'];
	return {
		innerShadowColor: context.parseColor(node),
		innerShadowOpacity: context.extractColorOpacity(node),
		innerShadowBlur: nonNegativePx(integer(effect.blurRadiusEmu), context.emuPerPx),
		innerShadowOffsetX: offset.x,
		innerShadowOffsetY: offset.y,
		innerShadowRotateWithShape:
			typeof rotateWithShape === 'boolean'
				? rotateWithShape
				: rotateWithShape === '1' || rotateWithShape === 'true'
					? true
					: rotateWithShape === '0' || rotateWithShape === 'false'
						? false
						: undefined,
	};
}

/** `a:glow` (20.1.8.32). */
export function glowStyle(
	effect: Effect<'glow'>,
	node: XmlObject,
	context: EffectStyleContext,
): Partial<ShapeStyle> {
	return {
		glowColor: context.parseColor(node),
		glowRadius: nonNegativePx(integer(effect.radiusEmu), context.emuPerPx),
		glowOpacity: context.extractColorOpacity(node),
	};
}

/** `a:softEdge` (20.1.8.53). */
export function softEdgeStyle(
	effect: Effect<'softEdge'>,
	context: EffectStyleContext,
): Partial<ShapeStyle> {
	return { softEdgeRadius: nonNegativePx(integer(effect.radiusEmu), context.emuPerPx) };
}

/** `a:blur` (20.1.8.15). Only a true `grow` is recorded. */
export function blurStyle(
	effect: Effect<'blur'>,
	context: EffectStyleContext,
): Partial<ShapeStyle> {
	return {
		blurRadius: nonNegativePx(integer(effect.radiusEmu), context.emuPerPx),
		blurGrow: effect.grow === true ? true : undefined,
	};
}

const fixedFraction = (value: number | undefined) => {
	const parsed = integer(value);
	return parsed !== undefined && parsed >= 0 && parsed <= 100000 ? parsed / 100000 : undefined;
};
const degrees = (value: number | undefined) =>
	value === undefined ? undefined : integer(value)! / 60000;

/** `a:reflection` (20.1.8.50). `rot` is outside the schema but kept when written. */
export function reflectionStyle(
	effect: Effect<'reflection'>,
	node: XmlObject,
	emuPerPx: number,
): Partial<ShapeStyle> {
	const rotation = Number.parseInt(String(node['@_rot'] ?? ''), 10);
	return {
		reflectionBlurRadius: nonNegativePx(integer(effect.blurRadiusEmu), emuPerPx),
		reflectionStartOpacity: fixedFraction(effect.startAlpha),
		reflectionEndOpacity: fixedFraction(effect.endAlpha),
		reflectionStartPosition: fixedFraction(effect.startPosition),
		reflectionEndPosition: fixedFraction(effect.endPosition),
		reflectionDirection: degrees(effect.directionAngle),
		reflectionRotation: Number.isFinite(rotation) ? rotation / 60000 : undefined,
		reflectionDistance: nonNegativePx(integer(effect.distanceEmu), emuPerPx),
		reflectionFadeDirection: degrees(effect.fadeDirectionAngle),
		reflectionScaleX: integer(effect.scaleX),
		reflectionScaleY: integer(effect.scaleY),
		reflectionSkewX: integer(effect.skewXAngle),
		reflectionSkewY: integer(effect.skewYAngle),
		reflectionAlignment: alignment<ShapeStyle['reflectionAlignment']>(effect.alignment),
		reflectionRotateWithShape: effect.rotateWithShape,
	};
}
