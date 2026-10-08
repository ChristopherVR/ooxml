import {
	clampUnitInterval,
	hslToRgb,
	parseDrawingFraction,
	parseDrawingHueDegrees,
	rgbToHsl,
} from '../color/index';
import { linearToSrgb255, srgb255ToLinear } from '../color/color-linear';
import type { DrawingColor } from './types';

/** Apply transforms in document order, keeping fractional channels until the final quantization. */
export function orderedColorTransforms(
	channels: { r: number; g: number; b: number },
	transforms: DrawingColor['transforms'],
): { r: number; g: number; b: number; alpha: number } {
	let { r, g, b } = channels;
	let alpha = 1;
	const wrap = (value: number) => ((value % 360) + 360) % 360;
	for (const transform of transforms) {
		const { name, value } = transform;
		const fraction = parseDrawingFraction(value);
		if (name === 'alpha' && fraction !== undefined) alpha = clampUnitInterval(fraction);
		else if (name === 'alphaMod' && fraction !== undefined)
			alpha = clampUnitInterval(alpha * fraction);
		else if (name === 'alphaOff' && fraction !== undefined)
			alpha = clampUnitInterval(alpha + fraction);
		else if (name === 'inv') [r, g, b] = [255 - r, 255 - g, 255 - b];
		else if (name === 'gray') r = g = b = 0.299 * r + 0.587 * g + 0.114 * b;
		else if ((name === 'shade' || name === 'tint') && fraction !== undefined) {
			const mix = (channel: number) =>
				linearToSrgb255(
					name === 'shade'
						? srgb255ToLinear(channel) * fraction
						: 1 - (1 - srgb255ToLinear(channel)) * fraction,
				);
			[r, g, b] = [mix(r), mix(g), mix(b)];
		} else {
			const hsl = rgbToHsl(r, g, b);
			if (name === 'comp') hsl.h = wrap(hsl.h + 180);
			else if (name === 'hue') {
				const hue = parseDrawingHueDegrees(value);
				if (hue !== undefined) hsl.h = wrap(hue);
			} else if (name === 'hueOff') {
				const hue = parseDrawingHueDegrees(value);
				if (hue !== undefined) hsl.h = wrap(hsl.h + hue);
			} else if (name === 'hueMod' && fraction !== undefined) hsl.h = wrap(hsl.h * fraction);
			else if (name === 'sat' && fraction !== undefined) hsl.s = clampUnitInterval(fraction);
			else if (name === 'satMod' && fraction !== undefined) hsl.s *= fraction;
			else if (name === 'satOff' && fraction !== undefined)
				hsl.s = clampUnitInterval(hsl.s + fraction);
			else if (name === 'lum' && fraction !== undefined) hsl.l = clampUnitInterval(fraction);
			else if (name === 'lumMod' && fraction !== undefined)
				hsl.l = clampUnitInterval(hsl.l * fraction);
			else if (name === 'lumOff' && fraction !== undefined)
				hsl.l = clampUnitInterval(hsl.l + fraction);
			else continue;
			({ r, g, b } = hslToRgb(hsl.h, hsl.s, hsl.l, (channel) => channel, {
				clampSaturation: false,
			}));
		}
	}
	return { r, g, b, alpha };
}
