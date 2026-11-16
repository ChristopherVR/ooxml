// The CT_EffectList effect reader (ECMA-376 Part 1, 20.1.8), written once over a DrawingNodeView
// so the DOM codec (`drawing-effects.ts`) and the pptx object-tree effect codecs read an effect's
// typed fields the same way. Source XML is the caller's: a DOM reader serializes the element.
import type { DrawingEffect } from './effect-model';
import { drawingBoolean } from './gradient-read';
import type { DrawingNodeView } from './node-view';
import type { DrawingFill } from './types';

/** Effect element local names and their typed kinds. */
export const DRAWING_EFFECT_KINDS = {
	outerShdw: 'outerShadow',
	innerShdw: 'innerShadow',
	prstShdw: 'presetShadow',
	glow: 'glow',
	softEdge: 'softEdge',
	blur: 'blur',
	reflection: 'reflection',
	fillOverlay: 'fillOverlay',
} as const;

export type DrawingEffectKind = Exclude<DrawingEffect['kind'], 'unsupported'>;
/** A list effect without its source XML. */
export type DrawingEffectFields = Exclude<DrawingEffect, { kind: 'unsupported' }>;

const GEOMETRY = { blurRadiusEmu: 'blurRad', distanceEmu: 'dist', directionAngle: 'dir' };
const TRANSFORM = { scaleX: 'sx', scaleY: 'sy', skewXAngle: 'kx', skewYAngle: 'ky' };

/** Numeric fields and their attributes, per kind (raw OOXML units). */
export const DRAWING_EFFECT_NUMBERS: Record<DrawingEffectKind, Readonly<Record<string, string>>> = {
	outerShadow: { ...GEOMETRY, ...TRANSFORM },
	innerShadow: GEOMETRY,
	presetShadow: { distanceEmu: 'dist', directionAngle: 'dir' },
	glow: { radiusEmu: 'rad' },
	softEdge: { radiusEmu: 'rad' },
	blur: { radiusEmu: 'rad' },
	fillOverlay: {},
	reflection: {
		...GEOMETRY,
		...TRANSFORM,
		fadeDirectionAngle: 'fadeDir',
		startAlpha: 'stA',
		startPosition: 'stPos',
		endAlpha: 'endA',
		endPosition: 'endPos',
	},
};
export const DRAWING_EFFECT_STRINGS: Partial<
	Record<DrawingEffectKind, Readonly<Record<string, string>>>
> = {
	outerShadow: { alignment: 'algn' },
	reflection: { alignment: 'algn' },
	presetShadow: { preset: 'prst' },
	fillOverlay: { blend: 'blend' },
};
export const DRAWING_EFFECT_BOOLEANS: Partial<
	Record<DrawingEffectKind, Readonly<Record<string, string>>>
> = {
	outerShadow: { rotateWithShape: 'rotWithShape' },
	reflection: { rotateWithShape: 'rotWithShape' },
	blur: { grow: 'grow' },
};
/** Kinds that carry a colour choice. */
export const DRAWING_EFFECT_COLORED: ReadonlySet<DrawingEffectKind> = new Set<DrawingEffectKind>([
	'outerShadow',
	'innerShadow',
	'presetShadow',
	'glow',
]);

/** The typed kind of an effect element's local name, if it is a list effect. */
export function drawingEffectKind(localName: string): DrawingEffectKind | undefined {
	return DRAWING_EFFECT_KINDS[localName as keyof typeof DRAWING_EFFECT_KINDS];
}

/**
 * Reads one list effect's typed fields; `undefined` when `localName` is not a list effect. No
 * defaults are synthesized. `readFill` reads the fill an `a:fillOverlay` holds, when the caller
 * can (the DOM fill reader); without it the overlay has no `fill`.
 */
export function readDrawingEffect<N>(
	view: DrawingNodeView<N>,
	localName: string,
	node: N,
	readFill?: (node: N) => DrawingFill | undefined,
): DrawingEffectFields | undefined {
	const kind = drawingEffectKind(localName);
	if (!kind) return undefined;
	const result: Record<string, unknown> = { kind };
	for (const [key, attribute] of Object.entries(DRAWING_EFFECT_NUMBERS[kind])) {
		const raw = view.attribute(node, attribute);
		if (raw === undefined || raw.trim() === '') continue;
		const value = Number(raw);
		if (Number.isFinite(value)) result[key] = value;
	}
	for (const [key, attribute] of Object.entries(DRAWING_EFFECT_STRINGS[kind] ?? {})) {
		const value = view.attribute(node, attribute);
		if (value !== undefined) result[key] = value;
	}
	for (const [key, attribute] of Object.entries(DRAWING_EFFECT_BOOLEANS[kind] ?? {})) {
		const value = drawingBoolean(view.attribute(node, attribute));
		if (value !== undefined) result[key] = value;
	}
	if (DRAWING_EFFECT_COLORED.has(kind)) {
		const color = view.color(node);
		if (color) result.color = color;
	}
	if (kind === 'fillOverlay' && readFill) {
		const fill = readFill(node);
		if (fill) result.fill = fill;
	}
	return result as unknown as DrawingEffectFields;
}

/** One child of an effect list: its typed fields (when it is a list effect) and its node. */
export interface DrawingEffectEntry<N> {
	name: string;
	node: N;
	effect?: DrawingEffectFields;
}

/** Reads every child of an effect list (or effect container), in the order the view gives. */
export function readDrawingEffectEntries<N>(
	view: DrawingNodeView<N>,
	list: N,
	readFill?: (node: N) => DrawingFill | undefined,
): DrawingEffectEntry<N>[] {
	return view.elements(list).map(({ name, node }) => {
		const effect = readDrawingEffect(view, name, node, readFill);
		return effect ? { name, node, effect } : { name, node };
	});
}
