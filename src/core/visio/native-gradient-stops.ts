import type { VisioGradientPaint, VisioFillGradient } from './model';
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

/** Native orthogonal stroke paint includes the physical half-width on each side. */
export function visioStrokeGradient(
	source: VisioFillGradient | undefined,
	lineWidth: number,
): VisioFillGradient | undefined {
	let gradient = source;
	if (gradient?.type === 'linear' && gradient.boundingBoxAngle === undefined) {
		// Native stroke paint spans the stroke's outer bounds, including its physical half-width.
		const dx = gradient.end[0] - gradient.start[0],
			dy = gradient.end[1] - gradient.start[1],
			length = Math.hypot(dx, dy),
			margin = lineWidth / 2;
		if (length > 0)
			gradient = {
				...gradient,
				start: [
					gradient.start[0] - (dx / length) * margin,
					gradient.start[1] - (dy / length) * margin,
				],
				end: [gradient.end[0] + (dx / length) * margin, gradient.end[1] + (dy / length) * margin],
			};
	}
	return gradient;
}
