/**
 * The copies `morph-text-body` paints a shape's body and text with, and the
 * keyframes each copy travels on.
 *
 * @module render/morph-text-body-copies
 */
import type { PptxElement } from 'ooxml-core/pptx';
import { hasShapeProperties } from 'ooxml-core/pptx';

import type { MorphTextTransform } from './morph-text-transform';
import { morphTextTail } from './morph-text-transform';

/** Style keys of an effect a copy could not paint once and only once. */
const EFFECT_KEYS =
	/^(shadow|innerShadow|glow|softEdge|reflection|effectList|effectDag|outerShadow|presetShadow|blur|scene3d|shape3d|dag|lineShadow|lineGlow|shapeFillOverlay|fillOverlay|inheritedEffectStyle)/u;
/** Effects only the body casts once the shape is filled: they sit under it. */
const UNDER_BODY_KEYS = /^(shadow|outerShadow|presetShadow)/u;
/** Style keys that paint the body: fills, outlines and their theme references. */
const BODY_KEYS =
	/^(fill|stroke|useBackgroundFill|inheritedFillStyle|inheritedLineStyle|lnRef|compoundLine|lineAlignment|customDash)/u;

function styleOf(element: PptxElement): Record<string, unknown> {
	return (hasShapeProperties(element) ? (element.shapeStyle ?? {}) : {}) as Record<string, unknown>;
}

function present(value: unknown): boolean {
	return (
		value !== undefined &&
		value !== null &&
		!(typeof value === 'object' && Object.keys(value as object).length === 0)
	);
}

/** Whether a source `a:effectLst` holds nothing but outer shadows. */
function onlyOuterShadows(value: unknown): boolean {
	return (
		typeof value === 'object' &&
		value !== null &&
		Object.keys(value).every((key) => key === 'a:outerShdw')
	);
}

/** Whether `element` paints an opaque fill, which hides whatever is cast under it. */
function opaquelyFilled(element: PptxElement): boolean {
	const style = styleOf(element);
	return (
		style.fillMode !== 'none' &&
		(style.fillOpacity ?? 1) === 1 &&
		typeof style.fillColor === 'string' &&
		!style.fillGradient &&
		!style.fillPatternPreset &&
		!style.fillImageUrl
	);
}

/**
 * Whether `element` carries an effect its body and text copies could not
 * share out. An outer shadow under an opaque fill is the body's alone (the
 * shape casts it, and the fill hides any the text would cast); measured on a
 * shadowed rounded rectangle whose text changes Calibri 28 pt -> Arial 40 pt,
 * PowerPoint 16 keeps the shadow under the body on every frame and carries
 * the text glyph over glyph with no shadow of its own. Every other effect
 * keeps the pair whole.
 */
export function morphCopiesCannotSplitEffects(element: PptxElement): boolean {
	const filled = opaquelyFilled(element);
	const blocks = (style: Record<string, unknown>): boolean =>
		Object.entries(style).some(([key, value]) => {
			if (key === 'inheritedEffectStyle') {
				// The theme's effect baseline: its effects are also on the style.
				return (
					typeof value === 'object' && value !== null && blocks(value as Record<string, unknown>)
				);
			}
			if (key === 'effectListXml' && filled && onlyOuterShadows(value)) {
				return false;
			}
			return EFFECT_KEYS.test(key) && present(value) && !(filled && UNDER_BODY_KEYS.test(key));
		});
	return blocks(styleOf(element));
}

/** `element` under `id` with its text taken out. */
export function morphBodyCopy(element: PptxElement, id: string): PptxElement {
	return { ...element, id, text: '', textSegments: [] } as PptxElement;
}

/**
 * `element` with its flips taken out, when its text reads the same without
 * them, else `undefined`. A flipped shape paints its text unmirrored
 * (`getTextCompensationTransform`), which is the text of the unflipped shape
 * as long as its insets mirror onto themselves and the compensation does not
 * swap axes (a box turned past 90 degrees and flipped one way only).
 */
export function morphUprightText(element: PptxElement): PptxElement | undefined {
	const { flipHorizontal: h, flipVertical: v } = element;
	if (!h && !v) {
		return element;
	}
	const rotation = (((element.rotation ?? 0) % 360) + 360) % 360;
	const ts = (element as { textStyle?: Record<string, number | undefined> }).textStyle ?? {};
	if (
		(h !== v && rotation > 90 && rotation < 270) ||
		(h && ts.bodyInsetLeft !== ts.bodyInsetRight) ||
		(v && ts.bodyInsetTop !== ts.bodyInsetBottom)
	) {
		return undefined;
	}
	return { ...element, flipHorizontal: false, flipVertical: false } as PptxElement;
}

/**
 * `element` under `id` with its fill, outline and body effects taken out,
 * painted unflipped (see {@link morphUprightText}).
 */
export function morphTextCopy(element: PptxElement, id: string): PptxElement {
	const kept = Object.fromEntries(
		Object.entries(styleOf(element)).filter(
			([key]) => !BODY_KEYS.test(key) && !EFFECT_KEYS.test(key),
		),
	);
	return {
		...element,
		id,
		flipHorizontal: false,
		flipVertical: false,
		shapeStyle: { ...kept, fillMode: 'none', strokeWidth: 0 },
	} as PptxElement;
}

function round(value: number): number {
	return Math.round(value * 1000) / 1000;
}

/**
 * One frame of a copy's journey: the transform, then its rotation, scale,
 * flips (stated only when the pair has any, so both frames share one
 * function list) and, for a turning text copy, the pivot.
 */
export function morphCopyTransform(
	t: MorphTextTransform,
	rotation: number,
	flips?: Pick<PptxElement, 'flipHorizontal' | 'flipVertical'>,
): string {
	const flipped = flips
		? ` scaleX(${flips.flipHorizontal ? -1 : 1}) scaleY(${flips.flipVertical ? -1 : 1})`
		: '';
	return `translate(${round(t.dx)}px, ${round(t.dy)}px) rotate(${round(rotation)}deg) scale(${round(t.sx)}, ${round(t.sy)})${flipped}${morphTextTail(t, flips ?? {})}`;
}

/** The box interpolation's incoming start: the incoming box laid over the outgoing one. */
export function morphBoxStart(from: PptxElement, to: PptxElement): MorphTextTransform {
	return {
		dx: from.x + from.width / 2 - (to.x + to.width / 2),
		dy: from.y + from.height / 2 - (to.y + to.height / 2),
		sx: Math.max(from.width, 1) / Math.max(to.width, 1),
		sy: Math.max(from.height, 1) / Math.max(to.height, 1),
	};
}

/** The box interpolation's ghost end: the outgoing box laid over the incoming one. */
export function morphBoxEnd(from: PptxElement, to: PptxElement): MorphTextTransform {
	return {
		dx: to.x + to.width / 2 - (from.x + from.width / 2),
		dy: to.y + to.height / 2 - (from.y + from.height / 2),
		sx: Math.max(to.width, 1) / Math.max(from.width, 1),
		sy: Math.max(to.height, 1) / Math.max(from.height, 1),
	};
}

/** Whether two transforms place the same box within a fraction of a pixel. */
export function sameMorphTransform(a: MorphTextTransform, b: MorphTextTransform): boolean {
	return (
		Math.abs(a.dx - b.dx) < 0.5 &&
		Math.abs(a.dy - b.dy) < 0.5 &&
		Math.abs(a.sx - b.sx) < 1e-3 &&
		Math.abs(a.sy - b.sy) < 1e-3 &&
		!a.tail &&
		!b.tail
	);
}

export function morphCopyKeyframes(
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
