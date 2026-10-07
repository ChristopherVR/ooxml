import type { VisioRadialGradient } from './model';

/** Shared preset centers, with native physical radii when bounds are supplied. */
export function radialFillGradient(
	direction: number,
	stops: VisioRadialGradient['stops'],
	bounds?: readonly [number, number],
): VisioRadialGradient | undefined {
	const presets: readonly (readonly [number, number, number])[] = [
		[1, 0, 1.4],
		[0, 0, 1.4],
		[0.5, 0.5, 0.73],
		[0.5, 1, 1.1],
		[0.5, 0, 1.1],
		[1, 1, 1.4],
		[0, 1, 1.4],
	];
	const preset = Number.isInteger(direction) ? presets[direction - 1] : undefined;
	if (preset && bounds) {
		const [width, height] = bounds;
		return {
			type: 'radial',
			coordinateSpace: 'local',
			center: [preset[0] * width, preset[1] * height],
			radius: Math.hypot(
				Math.max(preset[0], 1 - preset[0]) * width,
				Math.max(preset[1], 1 - preset[1]) * height,
			),
			stops,
		};
	}
	return preset
		? { type: 'radial', center: [preset[0], preset[1]], radius: preset[2], stops }
		: undefined;
}
