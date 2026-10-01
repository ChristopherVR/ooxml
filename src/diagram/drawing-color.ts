// DrawingML colour: a DOM reader and a theme-aware resolver. Written without any diagram
// knowledge so it moves unchanged into the future `drawingml` area (see docs/agnostic-core-plan.md).
import {
	clampUnitInterval,
	hexToRgbChannels,
	hslToRgb,
	parseDrawingFraction,
	parseDrawingHueDegrees,
	parseDrawingPercent,
	rgbToHsl,
	scrgbLinearToSrgb8,
	toHex,
} from '../color/index.js';
import { linearToSrgb255, srgb255ToLinear } from '../color/color-linear.js';
import { NS, elements, type XmlElement } from './dom.js';
import type { DiagramColor } from './types.js';

const COLOR_ELEMENTS: Record<string, DiagramColor['kind']> = {
	srgbClr: 'srgb',
	schemeClr: 'scheme',
	sysClr: 'system',
	prstClr: 'preset',
	scrgbClr: 'scrgb',
	hslClr: 'hsl',
};

function colorValue(kind: DiagramColor['kind'], element: XmlElement): string {
	const attribute = (name: string) => element.getAttribute(name) ?? '';
	if (kind === 'scrgb') return `${attribute('r')},${attribute('g')},${attribute('b')}`;
	if (kind === 'hsl') return `${attribute('hue')},${attribute('sat')},${attribute('lum')}`;
	return attribute('val');
}

/** Reads the colour element held directly by `container` (`a:solidFill`, `dgm:fillClrLst` item...). */
export function parseDrawingColorIn(container: XmlElement | undefined): DiagramColor | undefined {
	if (!container) return undefined;
	for (const element of elements(container)) {
		const kind = COLOR_ELEMENTS[element.localName];
		if (!kind || (element.namespaceURI && element.namespaceURI !== NS.a)) continue;
		return parseDrawingColor(element);
	}
	return undefined;
}

/** Reads one colour element (`a:srgbClr`, `a:schemeClr`, ...) and its transform children. */
export function parseDrawingColor(element: XmlElement): DiagramColor | undefined {
	const kind = COLOR_ELEMENTS[element.localName];
	if (!kind) return undefined;
	const color: DiagramColor = {
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
export function parseDrawingColorList(container: XmlElement | undefined): DiagramColor[] {
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

const PRESET_COLORS: Record<string, string> = {
	black: '#000000',
	white: '#FFFFFF',
	red: '#FF0000',
	green: '#008000',
	blue: '#0000FF',
	yellow: '#FFFF00',
	gray: '#808080',
	grey: '#808080',
};

function baseHex(color: DiagramColor, theme: DrawingColorTheme | undefined): string | undefined {
	switch (color.kind) {
		case 'srgb':
			return /^[0-9a-f]{6}$/i.test(color.value) ? `#${color.value.toUpperCase()}` : undefined;
		case 'scheme':
			return theme?.scheme(color.value);
		case 'system':
			return color.fallback && /^[0-9a-f]{6}$/i.test(color.fallback)
				? `#${color.fallback.toUpperCase()}`
				: undefined;
		case 'preset':
			return PRESET_COLORS[color.value];
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

const APPLIED = new Set([
	'alpha',
	'alphaMod',
	'alphaOff',
	'comp',
	'inv',
	'gray',
	'shade',
	'tint',
	'hue',
	'hueMod',
	'hueOff',
	'sat',
	'satMod',
	'satOff',
	'lum',
	'lumMod',
	'lumOff',
]);

/**
 * Resolves a colour to `#RRGGBB` plus alpha. Applies the transforms in the order PowerPoint does
 * (structural, shade/tint in linear light, one HSL round trip, alpha) and lists those it does not
 * apply (`red*`, `green*`, `blue*`, `gamma`) in `unapplied`. Returns `undefined` when the base
 * colour cannot be determined (an unknown theme slot, `phClr` with no host value, ...).
 */
export function resolveDrawingColor(
	color: DiagramColor,
	theme?: DrawingColorTheme,
): ResolvedDrawingColor | undefined {
	const base = baseHex(color, theme);
	const rgb = base ? hexToRgbChannels(base) : null;
	if (!rgb) return undefined;
	let { r, g, b } = rgb;
	const find = (name: string) => color.transforms.find((entry) => entry.name === name)?.value;
	const has = (name: string) => color.transforms.some((entry) => entry.name === name);

	if (has('comp')) {
		const hsl = rgbToHsl(r, g, b);
		({ r, g, b } = hslToRgb((hsl.h + 180) % 360, hsl.s, hsl.l));
	}
	if (has('inv')) [r, g, b] = [255 - r, 255 - g, 255 - b];
	if (has('gray')) r = g = b = Math.round(0.299 * r + 0.587 * g + 0.114 * b);

	const shade = parseDrawingPercent(find('shade'));
	if (shade !== undefined) {
		r = linearToSrgb255(srgb255ToLinear(r) * shade);
		g = linearToSrgb255(srgb255ToLinear(g) * shade);
		b = linearToSrgb255(srgb255ToLinear(b) * shade);
	}
	const tint = parseDrawingPercent(find('tint'));
	if (tint !== undefined) {
		r = linearToSrgb255(1 - (1 - srgb255ToLinear(r)) * tint);
		g = linearToSrgb255(1 - (1 - srgb255ToLinear(g)) * tint);
		b = linearToSrgb255(1 - (1 - srgb255ToLinear(b)) * tint);
	}

	if (['hue', 'hueMod', 'hueOff', 'sat', 'satMod', 'satOff', 'lum', 'lumMod', 'lumOff'].some(has)) {
		const hsl = rgbToHsl(r, g, b);
		const wrap = (degrees: number) => ((degrees % 360) + 360) % 360;
		const hueAbs = parseDrawingHueDegrees(find('hue'));
		if (hueAbs !== undefined) hsl.h = wrap(hueAbs);
		const hueMod = parseDrawingFraction(find('hueMod'));
		if (hueMod !== undefined) hsl.h = wrap(hsl.h * hueMod);
		const hueOff = parseDrawingHueDegrees(find('hueOff'));
		if (hueOff !== undefined) hsl.h = wrap(hsl.h + hueOff);
		const satAbs = parseDrawingFraction(find('sat'));
		if (satAbs !== undefined) hsl.s = clampUnitInterval(satAbs);
		const satMod = parseDrawingFraction(find('satMod'));
		if (satMod !== undefined) hsl.s = clampUnitInterval(hsl.s * satMod);
		const satOff = parseDrawingFraction(find('satOff'));
		if (satOff !== undefined) hsl.s = clampUnitInterval(hsl.s + satOff);
		const lumAbs = parseDrawingFraction(find('lum'));
		if (lumAbs !== undefined) hsl.l = clampUnitInterval(lumAbs);
		const lumMod = parseDrawingFraction(find('lumMod'));
		if (lumMod !== undefined) hsl.l = clampUnitInterval(hsl.l * lumMod);
		const lumOff = parseDrawingFraction(find('lumOff'));
		if (lumOff !== undefined) hsl.l = clampUnitInterval(hsl.l + lumOff);
		({ r, g, b } = hslToRgb(hsl.h, hsl.s, hsl.l));
	}

	let alpha = parseDrawingPercent(find('alpha')) ?? 1;
	const alphaMod = parseDrawingFraction(find('alphaMod'));
	if (alphaMod !== undefined) alpha = clampUnitInterval(alpha * alphaMod);
	const alphaOff = parseDrawingFraction(find('alphaOff'));
	if (alphaOff !== undefined) alpha = clampUnitInterval(alpha + alphaOff);

	return {
		hex: `#${toHex(r)}${toHex(g)}${toHex(b)}`,
		alpha,
		unapplied: [...new Set(color.transforms.map((entry) => entry.name))].filter(
			(name) => !APPLIED.has(name),
		),
	};
}
