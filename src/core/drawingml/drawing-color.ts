// DrawingML colour: a DOM reader and a theme-aware resolver (moved from the provisional copy in
// `diagram`, see docs/agnostic-core-plan.md).
import {
	hexToRgbChannels,
	hslToRgb,
	parseDrawingFraction,
	parseDrawingHueDegrees,
	scrgbLinearToSrgb8,
	toHex,
} from '../color/index';
import { APPLIED_COLOR_TRANSFORMS, applyDrawingColorTransformList } from './color-transforms';
import { DRAWING_PRESET_COLORS, DRAWING_SYSTEM_COLORS } from './preset-colors';
import { NS, elements, type XmlElement } from './dom';
import type { DrawingColor } from './types';
import { orderedColorTransforms } from './ordered-color-transforms';

const COLOR_ELEMENTS: Record<string, DrawingColor['kind']> = {
	srgbClr: 'srgb',
	schemeClr: 'scheme',
	sysClr: 'system',
	prstClr: 'preset',
	scrgbClr: 'scrgb',
	hslClr: 'hsl',
};

function colorValue(kind: DrawingColor['kind'], element: XmlElement): string {
	const attribute = (name: string) => element.getAttribute(name) ?? '';
	if (kind === 'scrgb') return `${attribute('r')},${attribute('g')},${attribute('b')}`;
	if (kind === 'hsl') return `${attribute('hue')},${attribute('sat')},${attribute('lum')}`;
	return attribute('val');
}

/** Reads the colour element held directly by `container` (`a:solidFill`, `dgm:fillClrLst` item...). */
export function parseDrawingColorIn(container: XmlElement | undefined): DrawingColor | undefined {
	if (!container) return undefined;
	for (const element of elements(container)) {
		const kind = COLOR_ELEMENTS[element.localName];
		if (!kind || (element.namespaceURI && element.namespaceURI !== NS.a)) continue;
		return parseDrawingColor(element);
	}
	return undefined;
}

/** Reads one colour element (`a:srgbClr`, `a:schemeClr`, ...) and its transform children. */
export function parseDrawingColor(element: XmlElement): DrawingColor | undefined {
	const kind = COLOR_ELEMENTS[element.localName];
	if (!kind) return undefined;
	const color: DrawingColor = {
		kind,
		value: colorValue(kind, element),
		transforms: elements(element).map((child) => ({
			name: child.localName,
			value: child.getAttribute('val') ?? '',
		})),
	};
	const fallback = kind === 'system' ? element.getAttribute('lastClr') : null;
	if (fallback) color.fallback = fallback;
	return color;
}

/** Every colour element directly under `container`, in order (a `dgm:fillClrLst`). */
export function parseDrawingColorList(container: XmlElement | undefined): DrawingColor[] {
	if (!container) return [];
	return elements(container).flatMap((element) => {
		const color = parseDrawingColor(element);
		return color ? [color] : [];
	});
}

/** Theme lookup the host supplies: `accent1`, `dk1`, `lt1`, `tx1`, `bg1`... to `#RRGGBB`. */
export interface DrawingColorTheme {
	scheme(name: string): string | undefined;
}

export interface ResolvedDrawingColor {
	/** `#RRGGBB`. */
	hex: string;
	/** 0..1; 1 when the colour has no alpha transform. */
	alpha: number;
	/** Transforms this resolver does not apply, so a caller can report an approximation. */
	unapplied: string[];
}

function baseHex(color: DrawingColor, theme: DrawingColorTheme | undefined): string | undefined {
	switch (color.kind) {
		case 'srgb':
			return /^[0-9a-f]{6}$/i.test(color.value) ? `#${color.value.toUpperCase()}` : undefined;
		case 'scheme':
			return theme?.scheme(color.value);
		case 'system':
			return color.fallback && /^[0-9a-f]{6}$/i.test(color.fallback)
				? `#${color.fallback.toUpperCase()}`
				: DRAWING_SYSTEM_COLORS[color.value];
		case 'preset':
			return DRAWING_PRESET_COLORS[color.value.toLowerCase()];
		case 'scrgb': {
			const [r, g, b] = color.value.split(',').map((part) => parseDrawingFraction(part));
			if (r === undefined || g === undefined || b === undefined) return undefined;
			return `#${toHex(scrgbLinearToSrgb8(r))}${toHex(scrgbLinearToSrgb8(g))}${toHex(scrgbLinearToSrgb8(b))}`;
		}
		case 'hsl': {
			const [h, s, l] = color.value.split(',');
			const hue = parseDrawingHueDegrees(h);
			const sat = parseDrawingFraction(s);
			const lum = parseDrawingFraction(l);
			if (hue === undefined || sat === undefined || lum === undefined) return undefined;
			const { r, g, b } = hslToRgb(hue, sat, lum);
			return `#${toHex(r)}${toHex(g)}${toHex(b)}`;
		}
	}
}

/** The document-order variant does not apply gamma or the RGB channel transforms. */
const NOT_IN_DOCUMENT_ORDER = /^(?:gamma|invGamma|red|green|blue)/u;

function unappliedTransforms(color: DrawingColor, documentOrder = false): string[] {
	return [...new Set(color.transforms.map((entry) => entry.name))].filter(
		(name) =>
			!APPLIED_COLOR_TRANSFORMS.has(name) || (documentOrder && NOT_IN_DOCUMENT_ORDER.test(name)),
	);
}

/**
 * Resolves a colour to `#RRGGBB` plus alpha. Applies the transforms in the order PowerPoint does
 * (`applyDrawingColorTransformList`: gamma, structural, shade/tint in linear light, one HSL round
 * trip, the RGB channels, alpha) and lists any it does not know in `unapplied`. Returns `undefined` when the base
 * colour cannot be determined (an unknown theme slot, `phClr` with no host value, ...).
 * `transformOrder: 'document'` applies repeated transforms in XML order with fractional
 * intermediate channels, matching measured native Excel gradient stops.
 */
export function resolveDrawingColor(
	color: DrawingColor,
	theme?: DrawingColorTheme,
	options?: { hslRounding?: 'nearest' | 'halfDown'; transformOrder?: 'legacy' | 'document' },
): ResolvedDrawingColor | undefined {
	const base = baseHex(color, theme);
	const rgb = base ? hexToRgbChannels(base) : null;
	if (!rgb) return undefined;
	if (options?.transformOrder === 'document') {
		const value = orderedColorTransforms(rgb, color.transforms);
		const round =
			options.hslRounding === 'halfDown'
				? (channel: number) => Math.ceil(channel - 0.5)
				: Math.round;
		return {
			hex: `#${toHex(round(value.r))}${toHex(round(value.g))}${toHex(round(value.b))}`,
			alpha: value.alpha,
			unapplied: unappliedTransforms(color, true),
		};
	}
	const value = applyDrawingColorTransformList(rgb, color.transforms, options);
	return {
		hex: `#${toHex(value.r)}${toHex(value.g)}${toHex(value.b)}`,
		alpha: value.alpha,
		unapplied: unappliedTransforms(color),
	};
}
