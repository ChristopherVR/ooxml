/**
 * The inspector's picture colour effects (PowerPoint's Picture Color options):
 * colour wash, recolour (`a:clrChange`) and duotone, each a toggleable section
 * of field descriptors every binding draws with its one generic renderer, and
 * {@link pictureDuotonePresetPatch}, the typed action of the duotone preset
 * swatches a binding keeps as a native template.
 *
 * The descriptors read a `PptxElement` and write a `Partial<PptxElement>`: the
 * picture's `imageEffects` with one effect changed, the others kept. A
 * non-picture reads every effect as off and writes nothing. The artistic
 * effect gallery stays a native template: two bindings draw it with picture
 * thumbnails.
 *
 * @module render/inspector/picture-effect-fields
 */
import type { PptxElement, PptxImageEffects } from 'ooxml-core/pptx';
import { isImageLikeElement } from 'ooxml-core/pptx';
import { normalizeHexColor } from 'ooxml-core/color';

import type { DuotonePreset } from '../image-duotone-canvas';
import type { InspectorSectionDescriptor } from './inspector-field-model';

export type PictureEffectSection = InspectorSectionDescriptor<PptxElement, Partial<PptxElement>>;

/** The colour wash a picture starts with when it is switched on. */
export const PICTURE_DEFAULT_COLOR_WASH = { color: '#0066cc', opacity: 40 } as const;
/** The recolour a picture starts with: white to black. */
export const PICTURE_DEFAULT_COLOR_CHANGE = { clrFrom: '#ffffff', clrTo: '#000000' } as const;
/** The duotone a picture starts with: black shadows, white highlights. */
export const PICTURE_DEFAULT_DUOTONE = { color1: '#000000', color2: '#ffffff' } as const;

const effectsOf = (element: PptxElement): PptxImageEffects | undefined =>
	isImageLikeElement(element) ? element.imageEffects : undefined;

/** `imageEffects` with `changes` merged in; nothing for a non-picture. */
function effectsPatch(
	element: PptxElement,
	changes: Partial<PptxImageEffects>,
): Partial<PptxElement> {
	if (!isImageLikeElement(element)) {
		return {};
	}
	return { imageEffects: { ...element.imageEffects, ...changes } } as Partial<PptxElement>;
}

const wash = (element: PptxElement) => effectsOf(element)?.colorWash ?? PICTURE_DEFAULT_COLOR_WASH;
const recolor = (element: PptxElement): NonNullable<PptxImageEffects['clrChange']> =>
	effectsOf(element)?.clrChange ?? PICTURE_DEFAULT_COLOR_CHANGE;
const duotone = (element: PptxElement) => effectsOf(element)?.duotone ?? PICTURE_DEFAULT_DUOTONE;

/** Colour wash (a tint laid over the picture): colour and strength. */
export const PICTURE_COLOR_WASH_SECTION: PictureEffectSection = {
	id: 'pictureColorWash',
	labelKey: 'pptx.image.colorWash',
	toggle: {
		testId: 'picture-colorWash-toggle',
		isOn: (element) => effectsOf(element)?.colorWash !== undefined,
		enable: (element) => effectsPatch(element, { colorWash: { ...PICTURE_DEFAULT_COLOR_WASH } }),
		disable: (element) => effectsPatch(element, { colorWash: undefined }),
	},
	fields: [
		{
			id: 'washColor',
			kind: 'color',
			labelKey: 'pptx.image.washColor',
			testId: 'picture-washColor',
			read: (element) => normalizeHexColor(wash(element).color, PICTURE_DEFAULT_COLOR_WASH.color),
			write: (element, value) =>
				effectsPatch(element, { colorWash: { ...wash(element), color: String(value) } }),
		},
		{
			id: 'washOpacity',
			kind: 'percent',
			labelKey: 'pptx.image.washOpacity',
			step: 1,
			unit: '%',
			decimals: 0,
			testId: 'picture-washOpacity',
			read: (element) => wash(element).opacity,
			write: (element, value) =>
				effectsPatch(element, { colorWash: { ...wash(element), opacity: Number(value) } }),
		},
	],
};

/** Recolour (`a:clrChange`): one colour of the picture replaced by another, or made transparent. */
export const PICTURE_RECOLOR_SECTION: PictureEffectSection = {
	id: 'pictureRecolor',
	labelKey: 'pptx.image.colorChange',
	toggle: {
		testId: 'picture-colorChange-toggle',
		isOn: (element) => effectsOf(element)?.clrChange !== undefined,
		enable: (element) =>
			effectsPatch(element, {
				clrChange: { ...PICTURE_DEFAULT_COLOR_CHANGE, clrToTransparent: false },
			}),
		disable: (element) => effectsPatch(element, { clrChange: undefined }),
	},
	fields: [
		{
			id: 'clrFrom',
			kind: 'color',
			labelKey: 'pptx.image.colorChangeFrom',
			testId: 'picture-colorChangeFrom',
			read: (element) =>
				normalizeHexColor(recolor(element).clrFrom, PICTURE_DEFAULT_COLOR_CHANGE.clrFrom),
			write: (element, value) =>
				effectsPatch(element, { clrChange: { ...recolor(element), clrFrom: String(value) } }),
		},
		{
			id: 'clrTo',
			kind: 'color',
			labelKey: 'pptx.image.colorChangeTo',
			testId: 'picture-colorChangeTo',
			visibleWhen: (element) => recolor(element).clrToTransparent !== true,
			read: (element) =>
				normalizeHexColor(recolor(element).clrTo, PICTURE_DEFAULT_COLOR_CHANGE.clrTo),
			write: (element, value) =>
				effectsPatch(element, { clrChange: { ...recolor(element), clrTo: String(value) } }),
		},
		{
			id: 'clrToTransparent',
			kind: 'toggle',
			labelKey: 'pptx.image.colorChangeTransparent',
			testId: 'picture-colorChangeTransparent',
			read: (element) => recolor(element).clrToTransparent === true,
			write: (element, value) =>
				effectsPatch(element, {
					clrChange: { ...recolor(element), clrToTransparent: value === true },
				}),
		},
	],
};

/** Duotone: the picture mapped between a shadow and a highlight colour. */
export const PICTURE_DUOTONE_SECTION: PictureEffectSection = {
	id: 'pictureDuotone',
	labelKey: 'pptx.image.duotone',
	toggle: {
		testId: 'picture-duotone-toggle',
		isOn: (element) => effectsOf(element)?.duotone !== undefined,
		enable: (element) => effectsPatch(element, { duotone: { ...PICTURE_DEFAULT_DUOTONE } }),
		disable: (element) => effectsPatch(element, { duotone: undefined }),
	},
	fields: [
		{
			id: 'duotoneShadow',
			kind: 'color',
			labelKey: 'pptx.image.duotoneShadows',
			testId: 'picture-duotoneShadow',
			read: (element) => normalizeHexColor(duotone(element).color1, PICTURE_DEFAULT_DUOTONE.color1),
			write: (element, value) =>
				effectsPatch(element, {
					duotone: { color1: String(value), color2: duotone(element).color2 },
				}),
		},
		{
			id: 'duotoneHighlight',
			kind: 'color',
			labelKey: 'pptx.image.duotoneHighlights',
			testId: 'picture-duotoneHighlight',
			read: (element) => normalizeHexColor(duotone(element).color2, PICTURE_DEFAULT_DUOTONE.color2),
			write: (element, value) =>
				effectsPatch(element, {
					duotone: { color1: duotone(element).color1, color2: String(value) },
				}),
		},
	],
};

/** The duotone preset swatches' action: apply `preset`'s two colours. */
export function pictureDuotonePresetPatch(
	element: PptxElement,
	preset: Pick<DuotonePreset, 'shadow' | 'highlight'>,
): Partial<PptxElement> {
	return effectsPatch(element, { duotone: { color1: preset.shadow, color2: preset.highlight } });
}

/** The three colour effects as a section list, what a generic renderer takes. */
export const PICTURE_EFFECT_SECTIONS: readonly PictureEffectSection[] = [
	PICTURE_COLOR_WASH_SECTION,
	PICTURE_RECOLOR_SECTION,
	PICTURE_DUOTONE_SECTION,
];
