import type { VisioFillGradient } from '../model';

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
	if (gradient.type === 'linear') {
		if (gradient.start.length !== 2 || gradient.end.length !== 2)
			throw new Error('The scene has invalid gradient endpoints.');
		for (const value of [...gradient.start, ...gradient.end])
			finite(value, 'gradient position', -20_000, 20_000);
	} else if (gradient.type === 'radial') {
		if (gradient.center.length !== 2) throw new Error('The scene has an invalid gradient center.');
		for (const value of gradient.center) finite(value, 'gradient center', 0, 1);
		finite(gradient.radius, 'gradient radius', Number.MIN_VALUE, 20_000);
	} else {
		if (!Array.isArray(gradient.regions) || ![2, 4].includes(gradient.regions.length))
			throw new Error('The scene has invalid gradient regions.');
		for (const region of gradient.regions) {
			finite(region.angle, 'gradient region angle', 0, 360);
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
				finite(value, 'gradient region endpoint', 0, 1);
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
	return gradient.stops.length * visioGradientInstances(gradient);
}
