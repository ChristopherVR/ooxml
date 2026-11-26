/**
 * The full built-in SmartArt colour and quick-style catalogue (`ooxml-core/diagram`) as the
 * PowerPoint galleries show it: dictionary keys per entry and section, the applied entry of a
 * diagram, and the patch a pick makes. A pick recolours the shapes in the editor through the
 * existing colour patch and nearest style intensity (an approximation), gives them the picked
 * quick style's 3D at once (`smartart-style-3d-preview.ts`), and records the built-in definition,
 * which save writes as the part so the reopened deck carries Office's own definition.
 *
 * @module render/ribbon-galleries/smartart-native-entries
 */
import {
	appliedDiagramNativeId,
	findDiagramNativeColor,
	findDiagramNativeStyle,
	DIAGRAM_NATIVE_COLOR_FOR_SCHEME,
	DIAGRAM_NATIVE_STYLE_FOR_INTENSITY,
	type DiagramNativeColorEntry,
	type DiagramNativeStyleEntry,
} from 'ooxml-core/diagram';
import type { PptxSmartArtData, SmartArtColorScheme, SmartArtStyle } from 'ooxml-core/pptx';

import { smartArtBuiltinLayouts } from '../smartart-builtin-layouts';

const COLORS = 'pptx.gallery.smartArtColors';
const STYLES = 'pptx.gallery.smartArtStyles';

const COLOR_VARIANTS = [
	'coloredOutline',
	'coloredFill',
	'gradientRange',
	'gradientLoop',
	'transparentGradientRange',
] as const;
const MAIN_SCHEME: Readonly<Record<string, string>> = {
	accent0_1: 'dark1Outline',
	accent0_2: 'dark2Outline',
	accent0_3: 'dark2Fill',
};
const STYLE_KEYS: Readonly<Record<string, string>> = {
	simple1: 'simpleFill',
	simple2: 'whiteOutline',
	simple3: 'subtleEffect',
	simple4: 'moderateEffect',
	simple5: 'intenseEffect',
	'3d1': 'polished',
	'3d2': 'inset',
	'3d3': 'cartoon',
	'3d4': 'powder',
	'3d5': 'brickScene',
	'3d6': 'flatScene',
	'3d7': 'metallicScene',
	'3d8': 'sunsetScene',
	'3d9': 'birdsEyeScene',
};

export interface SmartArtGalleryLabel {
	labelKey: string;
	labelParams?: Record<string, number>;
}

/** The dictionary key (and parameters) of a colour entry's name. */
export function smartArtColorLabel(entry: DiagramNativeColorEntry): SmartArtGalleryLabel {
	const main = MAIN_SCHEME[entry.id];
	if (main) return { labelKey: `${COLORS}.${main}` };
	if (entry.id === 'colorful1') return { labelKey: `${COLORS}.colorful1` };
	const range = /^colorful(\d)$/u.exec(entry.id);
	if (range)
		return {
			labelKey: `${COLORS}.colorfulRange`,
			labelParams: { from: Number(range[1]), to: Number(range[1]) + 1 },
		};
	const accent = /^accent(\d)_(\d)$/u.exec(entry.id);
	const variant = accent ? COLOR_VARIANTS[Number(accent[2]) - 1] : undefined;
	return variant
		? { labelKey: `${COLORS}.${variant}`, labelParams: { accent: Number(accent![1]) } }
		: { labelKey: `${COLORS}.${entry.id}` };
}

/** The dictionary key of a quick-style entry's name. */
export const smartArtStyleLabelKey = (entry: DiagramNativeStyleEntry): string =>
	`${STYLES}.${STYLE_KEYS[entry.id] ?? entry.id}`;

/** The dictionary key of a gallery section heading. */
export const smartArtColorSectionKey = (id: string): string => `${COLORS}.section.${id}`;
export const smartArtStyleSectionKey = (id: string): string =>
	`${STYLES}.section.${id === '3D' ? 'threeD' : 'bestMatch'}`;

const uniqueIdOf = (xml: string | undefined): string | undefined =>
	xml ? /\buniqueId="([^"]+)"/u.exec(xml)?.[1] : undefined;

/** The built-in colour entry a diagram carries: a pick, its loaded part, or its model scheme. */
export function appliedSmartArtColor(data: PptxSmartArtData): string | undefined {
	const modelId = data.colorScheme && DIAGRAM_NATIVE_COLOR_FOR_SCHEME[data.colorScheme];
	return (
		appliedDiagramNativeId('colors', uniqueIdOf(data.builtinColorsXml)) ??
		(data.colorTransformDirty ? modelId : undefined) ??
		appliedDiagramNativeId('colors', data.colorTransform?.uniqueId) ??
		modelId
	);
}

/** The built-in quick-style entry a diagram carries. */
export function appliedSmartArtStyle(data: PptxSmartArtData): string | undefined {
	return (
		appliedDiagramNativeId('style', uniqueIdOf(data.builtinQuickStyleXml)) ??
		appliedDiagramNativeId('style', data.quickStyle?.uniqueId) ??
		(data.style ? DIAGRAM_NATIVE_STYLE_FOR_INTENSITY[data.style] : undefined)
	);
}

/** The built-in entry of a gallery item id, also accepting a model scheme or intensity. */
export const nativeColorOf = (itemId: string) =>
	findDiagramNativeColor(appliedDiagramNativeId('colors', `colors/${itemId}`));
export const nativeStyleOf = (itemId: string) =>
	findDiagramNativeStyle(appliedDiagramNativeId('style', `style/${itemId}`));

/** The model scheme a built-in colour entry is, when it is one of the five. */
export function modelColorScheme(entry: DiagramNativeColorEntry): SmartArtColorScheme | undefined {
	return (Object.keys(DIAGRAM_NATIVE_COLOR_FOR_SCHEME) as SmartArtColorScheme[]).find(
		(scheme) => DIAGRAM_NATIVE_COLOR_FOR_SCHEME[scheme] === entry.id,
	);
}

/** The model intensity nearest a built-in quick style's effect strength. */
export function modelStyleIntensity(entry: DiagramNativeStyleEntry): SmartArtStyle {
	const exact = (Object.keys(DIAGRAM_NATIVE_STYLE_FOR_INTENSITY) as SmartArtStyle[]).find(
		(style) => DIAGRAM_NATIVE_STYLE_FOR_INTENSITY[style] === entry.id,
	);
	if (exact) return exact;
	return entry.preview.effect >= 3 ? 'intense' : entry.preview.effect >= 2 ? 'moderate' : 'flat';
}

/** The definition XML of a built-in entry, once the SmartArt library has loaded. */
export function builtinDefinitionXml(
	kind: 'colors' | 'style',
	uniqueId: string,
): string | undefined {
	const library = smartArtBuiltinLayouts();
	return library?.findBuiltinSmartArtStyleDefinition(kind, uniqueId)
		? library.loadBuiltinSmartArtStyleDefinitionXml(kind, uniqueId)
		: undefined;
}
