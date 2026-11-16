// The DrawingML colour transforms (ECMA-376 Part 1, 20.1.2.3) in the fixed order PowerPoint
// applies them, over 8-bit sRGB channels. Moved from the pptx colour codec so every format resolves
// a transform list the same way; `ordered-color-transforms.ts` is the document-order variant.
import {
	clampUnitInterval,
	hslToRgb,
	parseDrawingFraction,
	parseDrawingHueDegrees,
	parseDrawingPercent,
	rgbToHsl,
} from '../color/index';
import { linearToSrgb255, srgb255ToLinear } from '../color/color-linear';
import type { DrawingColor } from './types';

/** The transforms {@link applyDrawingColorTransformList} applies; any other name is reported. */
export const APPLIED_COLOR_TRANSFORMS: ReadonlySet<string> = new Set([
	'alpha',
	'alphaMod',
	'alphaOff',
	'gamma',
	'invGamma',
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
	'red',
	'redMod',
	'redOff',
	'green',
	'greenMod',
	'greenOff',
	'blue',
	'blueMod',
	'blueOff',
]);

const HSL_TRANSFORMS = [
	'hue',
	'hueMod',
	'hueOff',
	'sat',
	'satMod',
	'satOff',
	'lum',
	'lumMod',
	'lumOff',
];

/** IEC 61966-2-1 transfer function, linear to companded, one 8-bit channel. */
function gammaEncode(channel: number): number {
	const linear = clampUnitInterval(channel / 255);
	const companded = linear <= 0.0031308 ? 12.92 * linear : 1.055 * linear ** (1 / 2.4) - 0.055;
	return Math.min(255, Math.max(0, Math.round(companded * 255)));
}

/** IEC 61966-2-1 transfer function, companded to linear, one 8-bit channel. */
function gammaDecode(channel: number): number {
	const companded = clampUnitInterval(channel / 255);
	const linear = companded <= 0.04045 ? companded / 12.92 : ((companded + 0.055) / 1.055) ** 2.4;
	return Math.min(255, Math.max(0, Math.round(linear * 255)));
}

const byte = (value: number) => Math.min(255, Math.max(0, Math.round(value)));

/** One RGB channel's absolute, modulation and offset transforms (`red`, `redMod`, `redOff`). */
function channelTransforms(
	channel: number,
	find: (name: string) => string | undefined,
	name: 'red' | 'green' | 'blue',
): number {
	let value = channel;
	const absolute = parseDrawingFraction(find(name));
	if (absolute !== undefined) value = Math.round(clampUnitInterval(absolute) * 255);
	const mod = parseDrawingFraction(find(`${name}Mod`));
	if (mod !== undefined) value = byte(value * mod);
	const off = parseDrawingFraction(find(`${name}Off`));
	if (off !== undefined) value = byte(value + 255 * off);
	return value;
}

/**
 * Applies a transform list to 8-bit sRGB channels in PowerPoint's fixed order: gamma, structural
 * (`comp`, `inv`, `gray`), `shade` and `tint` mixed in linear light, one HSL round trip, the RGB
 * channel transforms, then alpha. Each transform is read once (the first of its name).
 */
export function applyDrawingColorTransformList(
	channels: { r: number; g: number; b: number },
	transforms: DrawingColor['transforms'],
	options?: { hslRounding?: 'nearest' | 'halfDown' },
): { r: number; g: number; b: number; alpha: number } {
	let { r, g, b } = channels;
	const find = (name: string) => transforms.find((entry) => entry.name === name)?.value;
	const has = (name: string) => transforms.some((entry) => entry.name === name);

	if (has('gamma')) [r, g, b] = [gammaEncode(r), gammaEncode(g), gammaEncode(b)];
	if (has('invGamma')) [r, g, b] = [gammaDecode(r), gammaDecode(g), gammaDecode(b)];
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

	if (HSL_TRANSFORMS.some(has)) {
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
		// Excel chart text rounds half-channel HSL ties downward (native 50% gray is #7F7F7F).
		const round =
			options?.hslRounding === 'halfDown' ? (value: number) => Math.ceil(value - 0.5) : Math.round;
		({ r, g, b } = hslToRgb(hsl.h, hsl.s, hsl.l, round));
	}

	r = channelTransforms(r, find, 'red');
	g = channelTransforms(g, find, 'green');
	b = channelTransforms(b, find, 'blue');

	let alpha = parseDrawingPercent(find('alpha')) ?? 1;
	const alphaMod = parseDrawingFraction(find('alphaMod'));
	if (alphaMod !== undefined) alpha = clampUnitInterval(alpha * alphaMod);
	const alphaOff = parseDrawingFraction(find('alphaOff'));
	if (alphaOff !== undefined) alpha = clampUnitInterval(alpha + alphaOff);
	return { r, g, b, alpha };
}
