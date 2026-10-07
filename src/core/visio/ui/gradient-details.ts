import type { VisioFillGradient } from '../model';
import {
	canUseVisioSigmaInterpolation,
	visioRenderedGradientStopCount,
} from '../native-gradient-stops';

/** Number of independently rendered gradients sharing the normalized stops. */
export function visioGradientInstances(gradient: VisioFillGradient | undefined): number {
	return gradient?.type === 'regions' ? gradient.regions.length : gradient ? 1 : 0;
}

/** Reuse the scene's finite-number checks and aggregate metadata accounting. */
export function assertVisioFillGradient(
	gradient: VisioFillGradient,
	finite: (value: number, label: string, minimum: number, maximum: number) => void,
	label: (value: string, limit: number) => void,
): number {
	if (
		!['linear', 'radial', 'regions'].includes(gradient.type) ||
		!Array.isArray(gradient.stops) ||
		gradient.stops.length < 2 ||
		gradient.stops.length > 128
	)
		throw new Error('The scene has an invalid fill gradient.');
	if (
		gradient.interpolation !== undefined &&
		(gradient.interpolation !== 'sigma-gamma22' || !canUseVisioSigmaInterpolation(gradient))
	)
		throw new Error('The scene has invalid native gradient interpolation.');
	if (gradient.type === 'linear') {
		if (gradient.boundingBoxAngle !== undefined)
			finite(gradient.boundingBoxAngle, 'gradient bounding-box angle', -360, 360);
		if (gradient.start.length !== 2 || gradient.end.length !== 2)
			throw new Error('The scene has invalid gradient endpoints.');
		for (const value of [...gradient.start, ...gradient.end])
			finite(
				value,
				'gradient position',
				gradient.boundingBoxAngle === undefined ? -20_000 : 0,
				gradient.boundingBoxAngle === undefined ? 20_000 : 1,
			);
	} else if (gradient.type === 'radial') {
		if (gradient.coordinateSpace !== undefined && gradient.coordinateSpace !== 'local')
			throw new Error('The scene has invalid radial gradient coordinates.');
		if (gradient.center.length !== 2) throw new Error('The scene has an invalid gradient center.');
		for (const value of gradient.center)
			finite(value, 'gradient center', 0, gradient.coordinateSpace === 'local' ? 20_000 : 1);
		finite(gradient.radius, 'gradient radius', Number.MIN_VALUE, 20_000);
	} else {
		if (gradient.coordinateSpace !== undefined && gradient.coordinateSpace !== 'shape')
			throw new Error('The scene has invalid region gradient coordinates.');
		const shapeCoordinates = gradient.coordinateSpace === 'shape';
		if (
			!Array.isArray(gradient.regions) ||
			(shapeCoordinates
				? gradient.regions.length < 1 || gradient.regions.length > 256
				: ![2, 4].includes(gradient.regions.length))
		)
			throw new Error('The scene has invalid gradient regions.');
		for (const region of gradient.regions) {
			finite(region.angle, 'gradient region angle', 0, shapeCoordinates ? 0 : 360);
			if (
				!Array.isArray(region.points) ||
				region.points.length !== 3 ||
				region.start.length !== 2 ||
				region.end.length !== 2
			)
				throw new Error('The scene has invalid gradient vertices.');
			for (const point of region.points) {
				if (point.length !== 2) throw new Error('The scene has invalid gradient vertices.');
				for (const value of point) finite(value, 'gradient vertex', 0, 1);
			}
			for (const value of [...region.start, ...region.end])
				finite(
					value,
					'gradient region endpoint',
					shapeCoordinates ? -2 : 0,
					shapeCoordinates ? 2 : 1,
				);
		}
	}
	let offset = -1;
	for (const stop of gradient.stops) {
		finite(stop.offset, 'gradient stop', 0, 1);
		finite(stop.opacity, 'gradient opacity', 0, 1);
		label(stop.color, 256);
		if (stop.offset < offset) throw new Error('Gradient stops must be ordered.');
		offset = stop.offset;
	}
	return visioRenderedGradientStopCount(gradient) * visioGradientInstances(gradient);
}
