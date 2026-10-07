import type { VisioGradientPaint, VisioFillGradient } from './model';
import { hexToRgbChannels } from '../color/color-primitives';
import { sigmaGradientStops } from '../color/sigma-gradient-stops';
import { VISIO_SIGMA_FACTORS } from './native-gradient-curve';
import { linearGradientEndpoints } from './theme-gradient';

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
	return sigmaGradientStops(paint.stops)!;
}

/** Native raster stroke paint projects through physical bounds including line width. */
export function visioStrokeGradient(
	source: VisioFillGradient | undefined,
	lineWidth: number,
	size?: readonly [number, number],
): VisioFillGradient | undefined {
	let gradient = source;
	if (gradient?.type === 'linear' && gradient.boundingBoxAngle !== undefined && size) {
		const { boundingBoxAngle, ...physical } = gradient;
		const endpoints = linearGradientEndpoints(
			size[0] + lineWidth,
			size[1] + lineWidth,
			-boundingBoxAngle * 60_000,
		);
		return {
			...physical,
			start: [endpoints.start[0] - lineWidth / 2, endpoints.start[1] - lineWidth / 2],
			end: [endpoints.end[0] - lineWidth / 2, endpoints.end[1] - lineWidth / 2],
		};
	}
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
