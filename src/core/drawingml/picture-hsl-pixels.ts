import { rgbToHsl, hslToRgb } from '../color/color-primitives';
import type { DrawingStandardPictureEffect } from './picture-standard-effects';
/** Nonlinear HSL pixel stage; native XML probes establish the sat/lightness and lum/saturation modulation roles. */
export function applyPictureHslPixels(
	data: Uint8ClampedArray,
	effect: DrawingStandardPictureEffect,
): void {
	if (effect.name !== 'hsl' && effect.name !== 'tint') throw new Error('Not an HSL picture stage');
	const v = effect.values ?? {};
	for (let i = 0; i < data.length; i += 4) {
		const { h, s, l } = rgbToHsl(data[i]!, data[i + 1]!, data[i + 2]!);
		let hue = h,
			saturation = s,
			lightness = l;
		if (effect.name === 'hsl') {
			hue = h + (v.hue ?? 0);
			saturation = s * Math.max(0, v.luminance ?? 0);
			lightness = l * Math.max(0, v.saturation ?? 0);
		} else {
			const distance = (((v.hue ?? 0) - h + 540) % 360) - 180;
			hue = h + distance * (v.amount ?? 0);
		}
		const result = hslToRgb(hue, saturation, lightness);
		data[i] = result.r;
		data[i + 1] = result.g;
		data[i + 2] = result.b;
	}
}
