/**
 * The pptx object-tree adapter of the shared DrawingML colour resolver.
 *
 * The transforms (ECMA-376 Part 1, 20.1.2.3), the preset and system colour tables and the base
 * colour of every colour-choice kind are resolved in `drawingml` (`applyDrawingColorTransformList`,
 * `resolveDrawingColor`); this module only reads a fast-xml-parser colour node into the neutral
 * {@link DrawingColor}. See docs/agnostic-core-plan.md, step 3.
 *
 * The object tree groups same-name siblings into an array, so a transform written twice (two
 * `a:tint`) arrives as one array entry; it is read as its occurrences in order, and the shared
 * resolver applies the first, as the DOM reader does.
 *
 * @module color-transforms
 */
import { hexToRgbChannels, toHex } from '../../../color/color-primitives';
import {
	applyDrawingColorTransformList,
	declaredDrawingColorAlpha,
} from '../../../drawingml/color-transforms';
import { resolveDrawingColor, type DrawingColorTheme } from '../../../drawingml/drawing-color';
import type { DrawingColor } from '../../../drawingml/types';
import type { XmlObject } from '../types';

type ColorChoiceOrder = ReadonlyArray<readonly [string, DrawingColor['kind']]>;

/** The colour-choice elements, in the order the runtime's codec tries them. */
export const CODEC_COLOR_CHOICE_ORDER: ColorChoiceOrder = [
	['a:scrgbClr', 'scrgb'],
	['a:srgbClr', 'srgb'],
	['a:sysClr', 'system'],
	['a:prstClr', 'preset'],
	['a:schemeClr', 'scheme'],
	// The table-style alias of `a:schemeClr`; its `val` is a scheme token (or `phClr`).
	['a:styleClr', 'scheme'],
	['a:hslClr', 'hsl'],
];

/**
 * The order `parseDrawingColorChoice` (the default-scheme reader) tries them: a holder that writes
 * two colour elements resolves to the same one it always did.
 */
export const DEFAULT_SCHEME_COLOR_CHOICE_ORDER: ColorChoiceOrder = [
	['a:scrgbClr', 'scrgb'],
	['a:srgbClr', 'srgb'],
	['a:sysClr', 'system'],
	['a:schemeClr', 'scheme'],
	['a:hslClr', 'hsl'],
	['a:prstClr', 'preset'],
];

const text = (value: unknown): string =>
	value === undefined || value === null ? '' : String(value);

/** The transform children (`a:shade`, `a:lumMod`...) of a colour node, in document order. */
export function objectColorTransforms(colorNode: XmlObject): DrawingColor['transforms'] {
	const transforms: DrawingColor['transforms'] = [];
	for (const [key, value] of Object.entries(colorNode)) {
		if (!key.startsWith('a:')) continue;
		const name = key.slice(2);
		for (const entry of Array.isArray(value) ? value : [value]) {
			const val = entry && typeof entry === 'object' ? (entry as XmlObject)['@_val'] : undefined;
			transforms.push({ name, value: text(val) });
		}
	}
	return transforms;
}

/** One colour element of the object tree (`a:srgbClr`'s node) as a neutral colour. */
export function objectDrawingColor(kind: DrawingColor['kind'], node: XmlObject): DrawingColor {
	const attribute = (name: string) => text(node[`@_${name}`]);
	const value =
		kind === 'scrgb'
			? `${attribute('r')},${attribute('g')},${attribute('b')}`
			: kind === 'hsl'
				? `${attribute('hue')},${attribute('sat')},${attribute('lum')}`
				: attribute('val').trim();
	const color: DrawingColor = { kind, value, transforms: objectColorTransforms(node) };
	const fallback = kind === 'system' ? attribute('lastClr').trim() : '';
	if (fallback) color.fallback = fallback;
	return color;
}

/**
 * Resolves a colour-choice holder (`a:solidFill`, a `a:clrScheme` slot...) to `#RRGGBB`, trying
 * each colour kind it holds in `order`. `resolveScheme` receives the lower-cased scheme
 * token; `phClr` takes `placeholderColor` first.
 */
export function resolveObjectColorChoice(
	colorChoice: XmlObject | undefined,
	resolveScheme: (schemeKey: string) => string | undefined,
	placeholderColor?: string,
	order: ColorChoiceOrder = CODEC_COLOR_CHOICE_ORDER,
): string | undefined {
	if (!colorChoice) return undefined;
	const theme: DrawingColorTheme = {
		scheme(name) {
			const key = name.toLowerCase();
			if (!key) return undefined;
			return key === 'phclr' ? (placeholderColor ?? resolveScheme('phclr')) : resolveScheme(key);
		},
	};
	for (const [element, kind] of order) {
		const node = colorChoice[element] as XmlObject | undefined;
		if (!node) continue;
		const resolved = resolveDrawingColor(
			objectDrawingColor(kind, typeof node === 'object' ? node : {}),
			theme,
		);
		if (resolved) return resolved.hex;
		// An unresolvable scheme reference ends the search, as the pptx readers always did.
		if (kind === 'scheme') return undefined;
	}
	return undefined;
}

/**
 * The opacity a colour-choice holder declares: the alpha the shared resolver gives the first
 * colour element it holds in `order` (`a:alpha`, then `a:alphaMod`, then `a:alphaOff`, each
 * clamped), or `undefined` when that element has no alpha transform.
 */
export function resolveObjectColorOpacity(
	colorChoice: XmlObject | undefined,
	order: ColorChoiceOrder = CODEC_COLOR_CHOICE_ORDER,
): number | undefined {
	if (!colorChoice) return undefined;
	for (const [element] of order) {
		const node = colorChoice[element];
		if (!node) continue;
		return typeof node === 'object' && !Array.isArray(node)
			? declaredDrawingColorAlpha(objectColorTransforms(node as XmlObject))
			: undefined;
	}
	return undefined;
}

/**
 * Apply every DrawingML colour transform found on {@link colorNode} to {@link baseColor} and
 * return the resulting `#RRGGBB`; a base that is not `#RRGGBB` comes back unchanged.
 *
 * The order is the shared resolver's (`applyDrawingColorTransformList`): gamma, structural,
 * shade and tint in linear light, one HSL round trip, then the RGB channel transforms.
 */
export function applyDrawingColorTransforms(baseColor: string, colorNode: XmlObject): string {
	const rgb = hexToRgbChannels(baseColor);
	if (!rgb) return baseColor;
	const { r, g, b } = applyDrawingColorTransformList(rgb, objectColorTransforms(colorNode));
	return `#${toHex(r)}${toHex(g)}${toHex(b)}`;
}
