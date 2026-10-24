/**
 * The Text Effects sections of the inspector (Format Text Effects: text shadow,
 * glow and reflection) as field descriptors every binding draws with its one
 * generic renderer. Each is a toggleable section: switching it on writes the
 * defaults below, switching it off clears every field of that effect.
 *
 * The descriptors read the element's `TextStyle` and produce a
 * `Partial<TextStyle>`. These are RUN properties (`a:rPr/a:effectLst`) and a
 * run is painted from its own style, so apply the patch to the body and every
 * run: `applyTextStyleUpdate(element, patch, null)`, the whole-text update the
 * ribbon uses. `textStylePatch` alone changes only the body, which the slide
 * does not show for text whose runs carry their own formatting.
 *
 * The defaults and bounds are the ones React's and Svelte's hand-built panels
 * use (the labels React's), so the same gesture writes the same `a:effectLst`
 * in every binding. The run CSS these fields drive is `text-effects.ts`
 * (`buildTextShadowCss`, `buildTextGlowFilter`); reflection renders through
 * `reflection.ts`.
 *
 * @module render/inspector/text-effect-fields
 */
import type { TextStyle } from 'ooxml-core/pptx';

import { normalizeHexColor } from '../fill-style';
import type {
	InspectorFieldDescriptor,
	InspectorFieldValue,
	InspectorSectionDescriptor,
} from './inspector-field-model';

/** What the Text Effects descriptors edit: the element's text style (absent for none). */
export type TextEffectTarget = TextStyle | undefined;
export type TextEffectSection = InspectorSectionDescriptor<TextEffectTarget, Partial<TextStyle>>;
type TextEffectField = InspectorFieldDescriptor<TextEffectTarget, Partial<TextStyle>>;

/** What ticking Text Shadow writes: a soft black shadow down and to the right. */
const SHADOW_ON = {
	textShadowColor: '#000000',
	textShadowBlur: 4,
	textShadowOffsetX: 2,
	textShadowOffsetY: 2,
	textShadowOpacity: 0.5,
} satisfies Partial<TextStyle>;

/** What ticking Text Glow writes: PowerPoint's yellow glow. */
const GLOW_ON = {
	textGlowColor: '#ffff00',
	textGlowRadius: 6,
	textGlowOpacity: 0.6,
} satisfies Partial<TextStyle>;

/** What ticking Text Reflection writes: a half-opacity reflection fading out. */
const REFLECTION_ON = {
	textReflection: true,
	textReflectionBlur: 1,
	textReflectionStartOpacity: 0.5,
	textReflectionEndOpacity: 0,
	textReflectionOffset: 3,
} satisfies Partial<TextStyle>;

/** What a field shows while its key is unset: the value ticking its effect on writes. */
const ON_DEFAULTS = { ...SHADOW_ON, ...GLOW_ON, ...REFLECTION_ON };

/** The same keys, all cleared: unticking drops the whole effect, never a stray field. */
function cleared<K extends keyof TextStyle>(on: Pick<TextStyle, K>): Partial<TextStyle> {
	return Object.fromEntries(Object.keys(on).map((key) => [key, undefined])) as Partial<TextStyle>;
}

const num = (v: InspectorFieldValue): number => Number(v);
const fxId = (section: string, field: string): string => `text-fx-${section}-${field}`;

function colorField(section: string, key: 'textShadowColor' | 'textGlowColor'): TextEffectField {
	return {
		id: 'color',
		kind: 'color',
		labelKey: 'pptx.textEffects.color',
		testId: fxId(section, 'color'),
		read: (ts) => normalizeHexColor(ts?.[key], ON_DEFAULTS[key]),
		write: (_ts, v) => ({ [key]: String(v) }),
	};
}

/** A 0-1 opacity box (two decimals, 0.05 steps), as the hand-built panels showed it. */
function opacityField(
	section: string,
	id: string,
	labelKey: string,
	key:
		| 'textShadowOpacity'
		| 'textGlowOpacity'
		| 'textReflectionStartOpacity'
		| 'textReflectionEndOpacity',
): TextEffectField {
	return {
		id,
		kind: 'number',
		labelKey,
		min: 0,
		max: 1,
		step: 0.05,
		testId: fxId(section, id),
		read: (ts) => ts?.[key] ?? ON_DEFAULTS[key],
		write: (_ts, v) => ({ [key]: num(v) }),
	};
}

/** A whole-number px box clamped to `[min, max]`. */
function pxField(
	section: string,
	id: string,
	labelKey: string,
	key:
		| 'textShadowBlur'
		| 'textShadowOffsetX'
		| 'textShadowOffsetY'
		| 'textGlowRadius'
		| 'textReflectionOffset',
	[min, max]: readonly [number, number],
): TextEffectField {
	return {
		id,
		kind: 'length',
		labelKey,
		min,
		max,
		step: 1,
		decimals: 0,
		testId: fxId(section, id),
		read: (ts) => ts?.[key] ?? ON_DEFAULTS[key],
		write: (_ts, v) => ({ [key]: num(v) }),
	};
}

const TEXT_SHADOW_SECTION: TextEffectSection = {
	id: 'shadow',
	labelKey: 'pptx.textEffects.shadow',
	toggle: {
		testId: 'text-fx-shadow-toggle',
		isOn: (ts) =>
			Boolean(ts?.textShadowColor) ||
			(typeof ts?.textShadowBlur === 'number' && ts.textShadowBlur > 0),
		enable: () => ({ ...SHADOW_ON }),
		disable: () => cleared(SHADOW_ON),
	},
	fields: [
		colorField('shadow', 'textShadowColor'),
		opacityField('shadow', 'opacity', 'pptx.textEffects.opacity', 'textShadowOpacity'),
		pxField('shadow', 'blur', 'pptx.textEffects.blur', 'textShadowBlur', [0, 50]),
		pxField('shadow', 'offsetX', 'pptx.textEffects.offsetX', 'textShadowOffsetX', [-50, 50]),
		pxField('shadow', 'offsetY', 'pptx.textEffects.offsetY', 'textShadowOffsetY', [-50, 50]),
	],
};

const TEXT_GLOW_SECTION: TextEffectSection = {
	id: 'glow',
	labelKey: 'pptx.textEffects.glow',
	toggle: {
		testId: 'text-fx-glow-toggle',
		isOn: (ts) =>
			Boolean(ts?.textGlowColor) ||
			(typeof ts?.textGlowRadius === 'number' && ts.textGlowRadius > 0),
		enable: () => ({ ...GLOW_ON }),
		disable: () => cleared(GLOW_ON),
	},
	fields: [
		colorField('glow', 'textGlowColor'),
		opacityField('glow', 'opacity', 'pptx.textEffects.opacity', 'textGlowOpacity'),
		pxField('glow', 'size', 'pptx.textEffects.size', 'textGlowRadius', [1, 50]),
	],
};

const TEXT_REFLECTION_SECTION: TextEffectSection = {
	id: 'reflection',
	labelKey: 'pptx.textEffects.reflection',
	toggle: {
		testId: 'text-fx-reflection-toggle',
		isOn: (ts) => Boolean(ts?.textReflection),
		enable: () => ({ ...REFLECTION_ON }),
		disable: () => cleared(REFLECTION_ON),
	},
	fields: [
		{
			id: 'blur',
			kind: 'length',
			labelKey: 'pptx.textEffects.blur',
			min: 0,
			max: 20,
			step: 0.5,
			decimals: 1,
			testId: fxId('reflection', 'blur'),
			read: (ts) => ts?.textReflectionBlur ?? ON_DEFAULTS.textReflectionBlur,
			write: (_ts, v) => ({ textReflectionBlur: num(v) }),
		},
		pxField('reflection', 'offset', 'pptx.textEffects.offset', 'textReflectionOffset', [0, 20]),
		opacityField(
			'reflection',
			'startOpacity',
			'pptx.textEffects.startOpacity',
			'textReflectionStartOpacity',
		),
		opacityField(
			'reflection',
			'endOpacity',
			'pptx.textEffects.endOpacity',
			'textReflectionEndOpacity',
		),
	],
};

/** Text shadow, glow and reflection, in the order every pane lists them. */
export const TEXT_EFFECT_SECTIONS: readonly TextEffectSection[] = [
	TEXT_SHADOW_SECTION,
	TEXT_GLOW_SECTION,
	TEXT_REFLECTION_SECTION,
];
