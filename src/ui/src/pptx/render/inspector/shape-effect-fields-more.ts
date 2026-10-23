/**
 * The reflection, blur and top-bevel sections of the shape Effects inspector
 * (see `shape-effect-fields.ts` for the shadows, glow and soft edges).
 *
 * @module render/inspector/shape-effect-fields-more
 */
import type { ShapeStyle } from 'ooxml-core/pptx';

import { reflectionStateOf, reflectionStyle } from '../effects-helpers';
import { TEXT_3D_MAX_BEVEL_PT, text3dEmuToPt, text3dPtToEmu } from '../text-3d-fields';
import { BEVEL_PRESETS } from '../text-3d-presets';
import type { InspectorFieldOption } from './inspector-field-model';
import { fx, num } from './shape-effect-fields';
import type { ShapeEffectField, ShapeEffectSection } from './shape-effect-fields';

/** Blur radius the Blur toggle applies, in px. */
const DEFAULT_BLUR = 4;
/** Top-bevel width and height the Bevel toggle applies (6pt), in EMU. */
const DEFAULT_BEVEL_EMU = 76200;

const reflectionField = (
	id: 'blurRadius' | 'startOpacity' | 'endOpacity' | 'distance' | 'direction',
	labelKey: string,
	max: number,
	testId: string,
): ShapeEffectField => ({
	id,
	kind: id === 'direction' ? 'angle' : 'number',
	labelKey,
	min: 0,
	max,
	testId: fx('reflection', testId),
	read: (s) => reflectionStateOf(s)[id],
	write: (s, v) => reflectionStyle({ ...reflectionStateOf(s), [id]: num(v) }),
});

export const REFLECTION: ShapeEffectSection = {
	id: 'reflection',
	labelKey: 'pptx.effects.reflection',
	toggle: {
		testId: 'fx-reflection-toggle',
		isOn: (s) => reflectionStateOf(s).enabled,
		enable: (s) => {
			const state = reflectionStateOf(s);
			// A reflection that reads as off has a zero start opacity; enable a visible one.
			return reflectionStyle({ ...state, startOpacity: state.startOpacity || 50 });
		},
		disable: () => ({
			reflectionBlurRadius: 0,
			reflectionStartOpacity: 0,
			reflectionEndOpacity: 0,
			reflectionDistance: 0,
			reflectionDirection: 0,
		}),
	},
	fields: [
		reflectionField('blurRadius', 'pptx.effects.blur', 20, 'blur'),
		reflectionField('startOpacity', 'pptx.effects.startPercent', 100, 'start'),
		reflectionField('endOpacity', 'pptx.effects.endPercent', 100, 'end'),
		reflectionField('distance', 'pptx.effects.distance', 50, 'distance'),
		reflectionField('direction', 'pptx.effects.direction', 359, 'direction'),
		{
			id: 'rotation',
			kind: 'angle',
			labelKey: 'pptx.arrange.rotation',
			min: 0,
			max: 359,
			testId: fx('reflection', 'rotation'),
			read: (s) => s?.reflectionRotation ?? 0,
			write: (_s, v) => ({ reflectionRotation: num(v) }),
		},
	],
};

export const BLUR: ShapeEffectSection = {
	id: 'blur',
	labelKey: 'pptx.effects.blur',
	toggle: {
		testId: 'fx-blur-toggle',
		isOn: (s) => (s?.blurRadius ?? 0) > 0,
		enable: (s) => ({ blurRadius: s?.blurRadius || DEFAULT_BLUR }),
		disable: () => ({ blurRadius: 0 }),
	},
	fields: [
		{
			id: 'radius',
			kind: 'length',
			labelKey: 'pptx.effects.radius',
			min: 0,
			max: 50,
			testId: fx('blur', 'radius'),
			read: (s) => s?.blurRadius ?? DEFAULT_BLUR,
			write: (_s, v) => ({ blurRadius: num(v) }),
		},
	],
};

/** The bevel presets a top bevel can take ("none" is the toggle being off). */
export const SHAPE_BEVEL_OPTIONS: readonly InspectorFieldOption[] = BEVEL_PRESETS.filter(
	(preset) => preset.value !== 'none',
).map((preset) => ({ value: preset.value, labelKey: preset.labelKey }));

const bevelSize = (id: 'bevelTopWidth' | 'bevelTopHeight', labelKey: string): ShapeEffectField => ({
	id,
	kind: 'length',
	labelKey,
	min: 0,
	max: TEXT_3D_MAX_BEVEL_PT,
	unit: 'pt',
	testId: fx('bevel', id === 'bevelTopWidth' ? 'width' : 'height'),
	read: (s) => text3dEmuToPt(s?.shape3d?.[id] ?? DEFAULT_BEVEL_EMU),
	write: (s, v) => ({ shape3d: { ...s?.shape3d, [id]: text3dPtToEmu(num(v)) } }),
});

export const BEVEL: ShapeEffectSection = {
	id: 'bevel',
	labelKey: 'pptx.text3d.topBevel',
	toggle: {
		testId: 'fx-bevel-toggle',
		isOn: (s) => Boolean(s?.shape3d?.bevelTopType) && s?.shape3d?.bevelTopType !== 'none',
		enable: (s) => ({
			shape3d: {
				...s?.shape3d,
				bevelTopType: s?.shape3d?.bevelTopType || 'circle',
				bevelTopWidth: s?.shape3d?.bevelTopWidth ?? DEFAULT_BEVEL_EMU,
				bevelTopHeight: s?.shape3d?.bevelTopHeight ?? DEFAULT_BEVEL_EMU,
			},
		}),
		disable: (s) => {
			const { bevelTopType: _t, bevelTopWidth: _w, bevelTopHeight: _h, ...rest } = s?.shape3d ?? {};
			return { shape3d: rest };
		},
	},
	fields: [
		{
			id: 'bevelTopType',
			kind: 'select',
			labelKey: 'pptx.text3d.bevelType',
			wide: true,
			options: SHAPE_BEVEL_OPTIONS,
			testId: fx('bevel', 'type'),
			read: (s) => s?.shape3d?.bevelTopType ?? 'circle',
			write: (s, v) => ({
				shape3d: {
					...s?.shape3d,
					bevelTopType: String(v) as NonNullable<ShapeStyle['shape3d']>['bevelTopType'],
				},
			}),
		},
		bevelSize('bevelTopWidth', 'pptx.text3d.bevelWidth'),
		bevelSize('bevelTopHeight', 'pptx.text3d.bevelHeight'),
	],
};
