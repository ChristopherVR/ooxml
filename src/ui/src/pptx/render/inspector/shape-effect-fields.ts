/**
 * The shape Effects sections of the inspector (Format Shape > Effects) as
 * field descriptors every binding draws with its one generic renderer: outer
 * shadow, inner shadow, glow and soft edges here, reflection, blur and the top
 * bevel in `shape-effect-fields-more.ts`, the full list in `SHAPE_EFFECT_SECTIONS`.
 *
 * The descriptors edit a `ShapeStyle` and produce a `Partial<ShapeStyle>`:
 * merge it into the element with `shapeQuickStylePatch` (bindings that patch
 * elements) or hand it to a shape-style update path as is. Reading and
 * clamping reuse the effect state helpers (`effects-helpers.ts`), so the
 * defaults a field shows are the ones the toggle writes.
 *
 * @module render/inspector/shape-effect-fields
 */
import type { ShapeStyle } from 'ooxml-core/pptx';

import { glowStateOf, glowStyle, softEdgeStateOf } from '../effects-helpers';
import {
	innerShadowStateOf,
	innerShadowStyle,
	outerShadowStateOf,
	outerShadowStyle,
} from '../effects-shadow-helpers';
import type { InnerShadowState, OuterShadowState } from '../effects-shadow-helpers';
import type {
	InspectorFieldDescriptor,
	InspectorFieldValue,
	InspectorSectionDescriptor,
} from './inspector-field-model';

/** What the Effects descriptors edit: the selected shape's style (absent for none). */
export type ShapeEffectTarget = ShapeStyle | undefined;
export type ShapeEffectField = InspectorFieldDescriptor<ShapeEffectTarget, Partial<ShapeStyle>>;
export type ShapeEffectSection = InspectorSectionDescriptor<ShapeEffectTarget, Partial<ShapeStyle>>;

/** PowerPoint's soft-edge radius when the toggle is ticked without one, in px. */
const DEFAULT_SOFT_EDGE = 6;

export const num = (v: InspectorFieldValue): number => Number(v);

const outer = (s: ShapeEffectTarget, changes: Partial<OuterShadowState>): Partial<ShapeStyle> =>
	outerShadowStyle({ ...outerShadowStateOf(s), ...changes });
const inner = (s: ShapeEffectTarget, changes: Partial<InnerShadowState>): Partial<ShapeStyle> =>
	innerShadowStyle({ ...innerShadowStateOf(s), ...changes });

/** The outer shadow's horizontal and vertical offsets from its angle and distance. */
function outerOffsets(s: ShapeEffectTarget): { x: number; y: number } {
	const { angle, distance } = outerShadowStateOf(s);
	const rad = (angle * Math.PI) / 180;
	return { x: Math.cos(rad) * distance, y: Math.sin(rad) * distance };
}

/** Angle (degrees, `[0, 360)`) and distance of an offset pair. */
function polar(x: number, y: number): { angle: number; distance: number } {
	return {
		angle: ((Math.atan2(y, x) * 180) / Math.PI + 360) % 360,
		distance: Math.sqrt(x * x + y * y),
	};
}

export const fx = (section: string, field: string): string => `fx-${section}-${field}`;

const colorField = (
	section: string,
	read: ShapeEffectField['read'],
	write: ShapeEffectField['write'],
): ShapeEffectField => ({
	id: 'color',
	kind: 'color',
	labelKey: 'pptx.effects.color',
	testId: fx(section, 'color'),
	read,
	write,
});

const opacityField = (
	section: string,
	read: ShapeEffectField['read'],
	write: ShapeEffectField['write'],
): ShapeEffectField => ({
	id: 'opacity',
	kind: 'percent',
	labelKey: 'pptx.inspector.opacity',
	unit: '%',
	testId: fx(section, 'opacity'),
	read,
	write,
});

export const OUTER_SHADOW: ShapeEffectSection = {
	id: 'outerShadow',
	labelKey: 'pptx.effects.outerShadow',
	toggle: {
		testId: 'fx-outer-shadow-toggle',
		isOn: (s) => outerShadowStateOf(s).enabled,
		enable: (s) => outerShadowStyle(outerShadowStateOf(s)),
		disable: () => ({ shadowColor: 'transparent' }),
	},
	fields: [
		colorField(
			'outer-shadow',
			(s) => outerShadowStateOf(s).color,
			(s, v) => outer(s, { color: String(v) }),
		),
		opacityField(
			'outer-shadow',
			(s) => outerShadowStateOf(s).opacity * 100,
			(s, v) => outer(s, { opacity: num(v) / 100 }),
		),
		{
			id: 'blur',
			kind: 'length',
			labelKey: 'pptx.effects.blur',
			min: 0,
			max: 96,
			testId: fx('outer-shadow', 'blur'),
			read: (s) => outerShadowStateOf(s).blur,
			write: (s, v) => outer(s, { blur: num(v) }),
		},
		{
			id: 'angle',
			kind: 'angle',
			labelKey: 'pptx.effects.angle',
			min: 0,
			max: 359,
			unit: '°',
			testId: fx('outer-shadow', 'angle'),
			read: (s) => outerShadowStateOf(s).angle,
			write: (s, v) => outer(s, { angle: num(v) }),
		},
		{
			id: 'distance',
			kind: 'length',
			labelKey: 'pptx.effects.distance',
			min: 0,
			max: 100,
			step: 0.5,
			testId: fx('outer-shadow', 'distance'),
			read: (s) => outerShadowStateOf(s).distance,
			write: (s, v) => outer(s, { distance: num(v) }),
		},
		{
			id: 'offsetX',
			kind: 'length',
			labelKey: 'pptx.effects.offsetX',
			min: -96,
			max: 96,
			testId: fx('outer-shadow', 'offset-x'),
			read: (s) => outerOffsets(s).x,
			write: (s, v) => outer(s, polar(num(v), outerOffsets(s).y)),
		},
		{
			id: 'offsetY',
			kind: 'length',
			labelKey: 'pptx.effects.offsetY',
			min: -96,
			max: 96,
			testId: fx('outer-shadow', 'offset-y'),
			read: (s) => outerOffsets(s).y,
			write: (s, v) => outer(s, polar(outerOffsets(s).x, num(v))),
		},
		{
			id: 'rotateWithShape',
			kind: 'toggle',
			labelKey: 'pptx.effects.rotateWithShape',
			wide: true,
			testId: 'fx-outer-shadow-rotate-with-shape',
			read: (s) => outerShadowStateOf(s).rotateWithShape,
			write: (s, v) => outer(s, { rotateWithShape: Boolean(v) }),
		},
	],
};

export const INNER_SHADOW: ShapeEffectSection = {
	id: 'innerShadow',
	labelKey: 'pptx.effects.innerShadow',
	toggle: {
		testId: 'fx-inner-shadow-toggle',
		isOn: (s) => innerShadowStateOf(s).enabled,
		enable: (s) => innerShadowStyle(innerShadowStateOf(s)),
		disable: () => ({ innerShadowColor: 'transparent' }),
	},
	fields: [
		colorField(
			'inner-shadow',
			(s) => innerShadowStateOf(s).color,
			(s, v) => inner(s, { color: String(v) }),
		),
		opacityField(
			'inner-shadow',
			(s) => innerShadowStateOf(s).opacity * 100,
			(s, v) => inner(s, { opacity: num(v) / 100 }),
		),
		{
			id: 'blur',
			kind: 'length',
			labelKey: 'pptx.effects.blur',
			min: 0,
			max: 96,
			testId: fx('inner-shadow', 'blur'),
			read: (s) => innerShadowStateOf(s).blur,
			write: (s, v) => inner(s, { blur: num(v) }),
		},
		{
			id: 'offsetX',
			kind: 'length',
			labelKey: 'pptx.effects.offsetX',
			min: -96,
			max: 96,
			testId: fx('inner-shadow', 'offset-x'),
			read: (s) => innerShadowStateOf(s).offsetX,
			write: (s, v) => inner(s, { offsetX: num(v) }),
		},
		{
			id: 'offsetY',
			kind: 'length',
			labelKey: 'pptx.effects.offsetY',
			min: -96,
			max: 96,
			testId: fx('inner-shadow', 'offset-y'),
			read: (s) => innerShadowStateOf(s).offsetY,
			write: (s, v) => inner(s, { offsetY: num(v) }),
		},
	],
};

export const GLOW: ShapeEffectSection = {
	id: 'glow',
	labelKey: 'pptx.effects.glow',
	toggle: {
		testId: 'fx-glow-toggle',
		isOn: (s) => glowStateOf(s).enabled,
		enable: (s) => glowStyle(glowStateOf(s)),
		disable: () => ({ glowColor: 'transparent', glowRadius: 0 }),
	},
	fields: [
		colorField(
			'glow',
			(s) => glowStateOf(s).color,
			(s, v) => glowStyle({ ...glowStateOf(s), color: String(v) }),
		),
		{
			id: 'radius',
			kind: 'length',
			labelKey: 'pptx.effects.radius',
			min: 0,
			max: 96,
			testId: fx('glow', 'radius'),
			read: (s) => glowStateOf(s).radius,
			write: (s, v) => glowStyle({ ...glowStateOf(s), radius: num(v) }),
		},
		opacityField(
			'glow',
			(s) => glowStateOf(s).opacity * 100,
			(s, v) => glowStyle({ ...glowStateOf(s), opacity: num(v) / 100 }),
		),
	],
};

export const SOFT_EDGE: ShapeEffectSection = {
	id: 'softEdge',
	labelKey: 'pptx.effects.softEdge',
	toggle: {
		testId: 'fx-soft-edge-toggle',
		isOn: (s) => softEdgeStateOf(s).enabled,
		enable: (s) => ({ softEdgeRadius: softEdgeStateOf(s).radius || DEFAULT_SOFT_EDGE }),
		disable: () => ({ softEdgeRadius: 0 }),
	},
	fields: [
		{
			id: 'radius',
			kind: 'length',
			labelKey: 'pptx.effects.radius',
			min: 0,
			max: 96,
			step: 0.5,
			testId: fx('soft-edge', 'radius'),
			read: (s) => softEdgeStateOf(s).radius,
			write: (_s, v) => ({ softEdgeRadius: num(v) }),
		},
	],
};
