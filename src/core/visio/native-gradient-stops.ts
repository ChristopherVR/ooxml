import type { VisioGradientPaint } from './model';
import { hexToRgbChannels, toHex } from '../color/color-primitives';
import { VISIO_SIGMA_FACTORS } from './native-gradient-curve';

export function canUseVisioSigmaInterpolation(paint: VisioGradientPaint): boolean {
	return (
		paint.stops.length === 2 &&
		paint.stops[0]?.offset === 0 &&
		paint.stops[1]?.offset === 1 &&
		paint.stops.every(
			(stop) =>
				stop.opacity === 1 && typeof stop.color === 'string' && !!hexToRgbChannels(stop.color),
		)
	);
}

/** Native sigma sampling expands paint only, never the saved ShapeSheet rows. */
export function visioRenderedGradientStopCount(paint: VisioGradientPaint): number {
	return paint.interpolation === 'sigma-gamma22' ? VISIO_SIGMA_FACTORS.length : paint.stops.length;
}

export function visioRenderedGradientStops(paint: VisioGradientPaint): VisioGradientPaint['stops'] {
	if (paint.interpolation !== 'sigma-gamma22') return paint.stops;
	if (!canUseVisioSigmaInterpolation(paint))
		throw new Error('The scene has invalid native gradient interpolation.');
	const front = hexToRgbChannels(paint.stops[0]!.color)!,
		back = hexToRgbChannels(paint.stops[1]!.color)!;
	return VISIO_SIGMA_FACTORS.map((value, index) => {
		const factor = value / 65536;
		const mix = (a: number, b: number) =>
			toHex(255 * ((a / 255) ** 2.2 * (1 - factor) + (b / 255) ** 2.2 * factor) ** (1 / 2.2));
		return {
			offset: Math.fround(index / 255),
			color: `#${mix(front.r, back.r)}${mix(front.g, back.g)}${mix(front.b, back.b)}`,
			opacity: 1,
		};
	});
}
